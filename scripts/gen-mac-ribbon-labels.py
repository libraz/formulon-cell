#!/usr/bin/env python3
"""Generate the Mac ribbon control labels in 7 locales from the installed Mac spreadsheet app.

Inputs (read-only; the app is never launched):
  - <app>/Contents/Resources/Excel.ribbonb: ribbon definition; supplies per-node
    `Label` redirects (`:rctcidstring(N)` takes TCID N's string).
  - <app>/Contents/Frameworks/mso40ui.framework: TCID -> string-key table embedded in the
    binary, and <loc>.lproj/{tbbtn,tbbtn2,tbfriendly}.strings.
  - scripts/mac-ribbon-label-tcids.json: our key (`labels:<k>`, `groups:<k>`, `menu:<id>`)
    -> TCID and the string source whose ja text equals the observed ja label.

Per entry and locale the label is the cleaned string of that source for the TCID's key; there
is no fallback between sources. Entries flagged `strip_ellipsis` additionally lose a trailing
ellipsis in every locale (see strip_trailing_ellipsis). Keys under `held` (unconfirmed source,
pinned pending, suspect value, no match) are not generated and keep their current text.

  gen-mac-ribbon-labels.py [OUT]        write the TS data file (default: the package's labels-bundle.ts)
  gen-mac-ribbon-labels.py --check PATH exit 1 if PATH differs from the generated output
"""
import argparse
import json
import plistlib
import re
import sys
from pathlib import Path

LOCS = {'en': 'en', 'ja': 'ja', 'de': 'de', 'fr': 'fr', 'ko': 'ko', 'th': 'th', 'zh': 'zh_CN'}
SOURCES = ('tbbtn', 'tbbtn2', 'tbfriendly')
KEY_TABLE_OFFSET = 7124152
KEY_TABLE_COUNT = 49990
NAMESPACES = ('labels', 'groups', 'menu')
REDIRECT = re.compile(r':rctcidstring\((\d+)\)')


LITERAL_AMP = '\0'


def clean(v):
    """Turn a command-table string (tbbtn, tbbtn2, tbfriendly) into the drawn label.

    Accelerator markers are dropped and a literal "&" is kept. Dialog nib strings get no such
    processing: the desktop draws them verbatim ("Find & Replace").

    Steps, in order; a literal "&" is held as a placeholder until the end so a later step
    cannot consume it:
      1. strip a "(&X)" or "(&&)" accelerator suffix;
      2. "&&" is a literal "&";
      3. a trailing "+&" is the key name "Ctrl+&";
      4. every other "&" is an accelerator marker and is removed.

    >>> clean('書式(&F)'), clean('クエリと接続(&&)')
    ('書式', 'クエリと接続')
    >>> clean('Find && Replace'), clean('คิวรี &&& &การเชื่อมต่อ')
    ('Find & Replace', 'คิวรี & การเชื่อมต่อ')
    >>> clean('Ctrl+&')
    'Ctrl+&'
    >>> clean('& Format Painter'), clean('Zum& ersten Bildschirm wechseln')
    ('Format Painter', 'Zum ersten Bildschirm wechseln')
    >>> clean('Afficher les modifications&')
    'Afficher les modifications'

    The ribbon treats a spaced " & " as a marker too and keeps both spaces (the ja Data tab
    draws "クエリ  接続").

    >>> clean('クエリ & 接続(&&)')
    'クエリ  接続'
    """
    v = re.sub(r'\s*\(&[^)]\)', '', v)
    v = v.replace('&&', LITERAL_AMP)
    v = re.sub(r'\+&$', '+' + LITERAL_AMP, v)
    v = v.replace('&', '').replace(LITERAL_AMP, '&')
    return v.replace('\u200b', '').strip()


def strip_trailing_ellipsis(v):
    """Drop a trailing "..." or "…" and the space before it.

    Rule, bundle side only: ribbon buttons are drawn without the ellipsis their bundle
    strings carry (every captured button agrees). A menu entry keeps its ellipsis unless
    the capture shows that menu item without one, so only map entries flagged
    `strip_ellipsis` use this.
    """
    return re.sub(r'\s*(\.\.\.|\u2026)$', '', v).strip()


def decode_ribbon(path):
    d = plistlib.load(open(path, 'rb'))
    objs = d['$objects']

    def r(x, depth=0):
        if isinstance(x, plistlib.UID):
            x = objs[x.data]
        if depth > 40:
            return None
        if isinstance(x, dict) and '$class' in x:
            c = objs[x['$class'].data]['$classname']
            if c in ('NSArray', 'NSMutableArray'):
                return [r(i, depth + 1) for i in x['NS.objects']]
            if c in ('NSDictionary', 'NSMutableDictionary'):
                return {str(r(k)): r(v, depth + 1) for k, v in zip(x['NS.keys'], x['NS.objects'])}
            if c in ('NSMutableString', 'NSString'):
                return x.get('NS.string')
            out = {'_c': c}
            for k, v in x.items():
                if k != '$class':
                    out[k] = r(v, depth + 1)
            return out
        return x

    return r(d['$top']['kOUIOfficeSpaceRibbonKey'])


def collect_redirects(tree):
    """TCID -> Label string, for every node that carries one."""
    found, conflicts = {}, []

    def walk(n):
        if isinstance(n, list):
            for v in n:
                walk(v)
        elif isinstance(n, dict):
            a = n.get('kOUIRibbonAttributesKey')
            if isinstance(a, dict) and isinstance(a.get('Tcid'), int) and isinstance(a.get('Label'), str):
                old = found.setdefault(a['Tcid'], a['Label'])
                if old != a['Label']:
                    conflicts.append((a['Tcid'], old, a['Label']))
            for k, v in n.items():
                if k != '_c':
                    walk(v)

    walk(tree)
    return found, conflicts


class Bundle:
    def __init__(self, app):
        base = Path(app) / 'Contents'
        fw = base / 'Frameworks/mso40ui.framework'
        raw = (fw / 'Versions/A/mso40ui').read_bytes()
        self.keys = [k.decode() for k in raw[KEY_TABLE_OFFSET:].split(b'\0')[:KEY_TABLE_COUNT]]
        self.strings = {
            loc: {s: plistlib.load(open(fw / 'Resources' / f'{d}.lproj/{s}.strings', 'rb')) for s in SOURCES}
            for loc, d in LOCS.items()
        }
        self.redirect, self.conflicts = collect_redirects(decode_ribbon(base / 'Resources/Excel.ribbonb'))

    def labels(self, tcid, source):
        """Return ({loc: text}, None) or (None, reason)."""
        src = tcid
        lab = self.redirect.get(tcid)
        if lab is not None:
            m = REDIRECT.fullmatch(lab)
            if not m:
                return None, f'TCID {tcid} Label {lab} points at an unknown string table'
            src = int(m.group(1))
        if src >= len(self.keys):
            return None, f'TCID {src} is outside the key table'
        key = self.keys[src]
        out = {}
        for loc in LOCS:
            v = self.strings[loc][source].get(key)
            if v is None:
                return None, f'TCID {src} (key {key}) has no {source} string for {loc}'
            out[loc] = clean(v)
        return out, None


def out_of_scope(ours):
    """Menu ids whose label is the id itself (gallery tiles) or an engine-localized function name (mac.function.*, mac.autosum.*)."""
    gallery, function = set(), set()
    for mid, v in ours['menu'].items():
        if mid.startswith(('mac.function.', 'mac.autosum.')):
            function.add(f'menu:{mid}')
        elif any(x == mid for x in v.values()):
            gallery.add(f'menu:{mid}')
    return gallery, function


def q(v):
    return "'" + v.replace('\\', '\\\\').replace("'", "\\'").replace('\n', '\\n') + "'"


def render(rows):
    out = [
        '// Generated by scripts/gen-mac-ribbon-labels.py from the Mac app bundle.',
        '// Do not edit by hand; edit scripts/mac-ribbon-label-tcids.json and regenerate.',
        '',
        'export const MAC_BUNDLE_LABELS: Readonly<',
        "  Record<string, Readonly<Record<'en' | 'ja' | 'de' | 'fr' | 'ko' | 'th' | 'zh', string>>>",
        '> = {',
    ]
    for key in sorted(rows):
        out.append(f'  {q(key)}: {{')
        out.extend(f'    {loc}: {q(rows[key][loc])},' for loc in LOCS)
        out.append('  },')
    out.append('};')
    return '\n'.join(out) + '\n'


def build(b, ours, mapfile):
    m = json.load(open(mapfile, encoding='utf-8'))
    known = {f'{ns}:{k}' for ns in NAMESPACES for k in ours[ns]} if ours else None
    rows, problems = {}, []
    for key, e in m['map'].items():
        if known is not None and key not in known:
            problems.append((key, 'key not present in the current label set'))
            continue
        if key in m.get('held', {}):
            problems.append((key, 'key is both mapped and held'))
            continue
        labs, why = b.labels(e['tcid'], e['source'])
        if labs is None:
            problems.append((key, why))
            continue
        if e.get('strip_ellipsis'):
            labs = {loc: strip_trailing_ellipsis(v) for loc, v in labs.items()}
        rows[key] = labs
    held = {k: v for k, v in m.get('held', {}).items()}
    return rows, problems, held


def report(rows, problems, held, ours):
    gallery, function = out_of_scope(ours)
    diff = {loc: 0 for loc in LOCS}
    for key, labs in rows.items():
        ns, _, name = key.partition(':')
        for loc in LOCS:
            if labs[loc] != clean(ours[ns][name][loc]):
                diff[loc] += 1
    total = sum(len(ours[ns]) for ns in NAMESPACES)
    by = {}
    for v in held.values():
        by[v['status']] = by.get(v['status'], 0) + 1
    w = lambda s: print(s, file=sys.stderr)  # noqa: E731
    w(f'keys {total}: generated {len(rows)}, held {len(held)} {by}, '
      f'out of scope {len(gallery)} gallery + {len(function)} function')
    w('generated labels that differ from the current text:')
    for loc in LOCS:
        w(f'  {loc}: {diff[loc]}')
    for key, why in problems:
        w(f'ERROR {key}: {why}')
    for key, v in sorted(held.items()):
        w(f"held [{v['status']}] {key}" + (f" (TCID {v['tcid']} {v.get('name', '')})" if v.get('tcid') else '')
          + f": {v['reason']}")


def main():
    p = argparse.ArgumentParser(description='Generate Mac ribbon labels from the app bundle.')
    p.add_argument(
        'out',
        nargs='?',
        default=str(
            Path(__file__).resolve().parent.parent
            / 'packages/formulon-cell/src/toolbar/ribbon/mac/labels-bundle.ts'
        ),
    )
    p.add_argument('--check', metavar='PATH')
    p.add_argument('--app', default='/Applications/Microsoft Excel.app')
    p.add_argument('--ours', help='current labels JSON {labels,groups,menu}; enables the diff report')
    p.add_argument('--map', default=str(Path(__file__).with_name('mac-ribbon-label-tcids.json')))
    a = p.parse_args()
    b = Bundle(a.app)
    for t, x, y in b.conflicts:
        print(f'warning: TCID {t} has conflicting Label values {x} / {y}; using the first', file=sys.stderr)
    ours = json.load(open(a.ours, encoding='utf-8')) if a.ours else None
    rows, problems, held = build(b, ours, a.map)
    if ours:
        report(rows, problems, held, ours)
    else:
        for key, why in problems:
            print(f'ERROR {key}: {why}', file=sys.stderr)
    text = render(rows)
    if problems:
        return 2
    if a.check:
        path = Path(a.check)
        if not path.exists() or path.read_text(encoding='utf-8') != text:
            print(f'{a.check} differs from the generated output', file=sys.stderr)
            return 1
        return 0
    Path(a.out).write_text(text, encoding='utf-8')
    return 0


if __name__ == '__main__':
    sys.exit(main())
