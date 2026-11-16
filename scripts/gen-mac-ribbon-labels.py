#!/usr/bin/env python3
"""Generate the Mac ribbon control labels in 7 locales from the installed Mac spreadsheet app.

Inputs (read-only; the app is never launched):
  - <app>/Contents/Resources/Excel.ribbonb: ribbon definition; supplies per-node
    `Label` redirects (`:rctcidstring(N)` takes TCID N's string).
  - <app>/Contents/Frameworks/mso40ui.framework: TCID -> string-key table embedded in the
    binary, and <loc>.lproj/{tbbtn,tbbtn2,tbfriendly}.strings.
  - <app>/Contents/**/<loc>.lproj/*.strings: any other strings table, named by a map entry.
  - scripts/mac-ribbon-label-tcids.json: our key -> its string source, either a TCID and the
    command table (tbbtn|tbbtn2|tbfriendly) whose ja text equals the observed ja label, or a
    strings file and key (see strings_path). Ribbon controls use the `labels:`, `groups:` and
    `menu:` namespaces; other namespaces name the menus, galleries and names read elsewhere.

Per entry and locale the label is the cleaned string of its source; there is no fallback
between sources. An entry with `identical_to: [tcid, source]` fails the run
when that candidate's strings differ in any locale from the mapped ones. Entries flagged
`strip_ellipsis` additionally lose a trailing
ellipsis in every locale (see strip_trailing_ellipsis). Keys under `held` (unconfirmed source,
pinned pending, suspect value, no match) are not generated and keep their current text, unless
the entry carries locale text: `text` ({locale: text}, the measured or current wording) and/or
`provisional` (a strings file and key whose ja differs or is unconfirmed). A held locale takes
`text`, else the provisional source, else `text.en`.

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
FRAMEWORK_SUFFIX = '.framework'
BUNDLE_SUFFIX = '.bundle'
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
      4. every other "&" is an accelerator marker and is removed; a leading one also takes the
         single space after it.

    >>> clean('書式(&F)'), clean('クエリと接続(&&)')
    ('書式', 'クエリと接続')
    >>> clean('Find && Replace'), clean('คิวรี &&& &การเชื่อมต่อ')
    ('Find & Replace', 'คิวรี & การเชื่อมต่อ')
    >>> clean('Ctrl+&')
    'Ctrl+&'
    >>> clean('& Format Painter'), clean('Zum& ersten Bildschirm wechseln')
    ('Format Painter', 'Zum ersten Bildschirm wechseln')
    >>> clean('表格  ')
    '表格  '
    >>> clean('Afficher les modifications&')
    'Afficher les modifications'

    Every other code point passes through as the bundle has it (the ribbon draws these
    zero-width spaces).

    >>> clean('次の\u200bコメント') == '次の\u200bコメント'
    True

    The ribbon treats a spaced " & " as a marker too and keeps both spaces (the ja Data tab
    draws "クエリ  接続").

    >>> clean('クエリ & 接続(&&)')
    'クエリ  接続'
    """
    v = re.sub(r'\s*\(&[^)]\)', '', v)
    v = v.replace('&&', LITERAL_AMP)
    v = re.sub(r'\+&$', '+' + LITERAL_AMP, v)
    v = re.sub(r'^& ?', '', v)
    return v.replace('&', '').replace(LITERAL_AMP, '&')


def strip_trailing_ellipsis(v):
    """Drop a trailing "..." or "…" and the space before it, also when zero-width spaces follow.

    Rule, bundle side only: ribbon buttons are drawn without the ellipsis their bundle
    strings carry (every captured button agrees). A menu entry keeps its ellipsis unless
    the capture shows that menu item without one, so only map entries flagged
    `strip_ellipsis` use this. Zero-width spaces after the ellipsis stay.

    >>> strip_trailing_ellipsis('슬라이서...\u200b\u200b') == '슬라이서\u200b\u200b'
    True
    >>> strip_trailing_ellipsis('Online Pictures ...'), strip_trailing_ellipsis('Gitter\u2026')
    ('Online Pictures', 'Gitter')
    """
    return re.sub(r'\s*(\.\.\.|\u2026)(?=\u200b*$)', '', v)


def strings_path(contents, file, lproj):
    """Path of a strings table in one locale.

    `file` is the table name, prefixed by the frameworks and bundles that hold it; a bare name
    is in the app's own Resources.

    >>> from pathlib import PurePosixPath as P
    >>> str(strings_path(P('C'), 'strman_0066.strings', 'de'))
    'C/Resources/de.lproj/strman_0066.strings'
    >>> str(strings_path(P('C'), 'mso99.framework/mso_intl.bundle/strman_ac16.strings', 'zh_CN'))
    'C/Frameworks/mso99.framework/Resources/mso_intl.bundle/Contents/Resources/zh_CN.lproj/strman_ac16.strings'
    """
    *holders, name = file.split('/')
    res = contents / 'Resources'
    for h in holders:
        if h.endswith(FRAMEWORK_SUFFIX):
            res = contents / 'Frameworks' / h / 'Resources'
        elif h.endswith(BUNDLE_SUFFIX):
            res = res / h / 'Contents' / 'Resources'
        else:
            raise ValueError(f'{file}: {h} is neither a framework nor a bundle')
    return res / f'{lproj}.lproj' / name


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
        self.contents = base
        self.tables = {}

    def table(self, file, loc):
        if (file, loc) not in self.tables:
            path = strings_path(self.contents, file, LOCS[loc])
            self.tables[(file, loc)] = plistlib.load(open(path, 'rb')) if path.exists() else None
        return self.tables[(file, loc)]

    def file_labels(self, file, key):
        """Return ({loc: text}, None) or (None, reason) for a strings file and key."""
        out = {}
        for loc in LOCS:
            t = self.table(file, loc)
            if t is None:
                return None, f'{file} has no {loc} table'
            v = t.get(key)
            if not isinstance(v, str):
                return None, f'{file} has no {loc} string for key {key}'
            out[loc] = clean(v)
        return out, None

    def source(self, spec):
        """Labels of a map entry's source: a strings file and key, or a TCID and command table."""
        if isinstance(spec, dict) and 'file' in spec:
            return self.file_labels(spec['file'], spec['key'])
        if isinstance(spec, dict):
            return self.labels(spec['tcid'], spec['source'])
        return self.labels(*spec)

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
    """Menu ids whose label is the id itself (gallery tiles) or an engine-localized function name (mac.function.*)."""
    gallery, function = set(), set()
    for mid, v in ours['menu'].items():
        if mid.startswith('mac.function.'):
            function.add(f'menu:{mid}')
        elif any(x == mid for x in v.values()):
            gallery.add(f'menu:{mid}')
    return gallery, function


def q(v):
    """A TS string literal quoted the way the formatter writes it: single quotes unless the
    value holds a single quote and no double quote."""
    v = v.replace('\\', '\\\\').replace('\n', '\\n')
    if "'" in v and '"' not in v:
        return '"' + v + '"'
    return "'" + v.replace("'", "\\'") + "'"


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
        if known is not None and key.partition(':')[0] in NAMESPACES and key not in known:
            problems.append((key, 'key not present in the current label set'))
            continue
        if key in m.get('held', {}):
            problems.append((key, 'key is both mapped and held'))
            continue
        if key.startswith('menu:') and e.get('evidence') and not e.get('ax_path'):
            # A drawn reading for a menu item must name the menu it was read from.
            problems.append((key, 'menu evidence without the AX path of the menu item'))
            continue
        control = e.get('control')
        if control == 'button' and not e.get('strip_ellipsis'):
            problems.append((key, 'a button is drawn without the bundle ellipsis: set strip_ellipsis'))
            continue
        if control == 'menu-item' and e.get('strip_ellipsis'):
            problems.append((key, 'a menu item keeps its ellipsis: remove strip_ellipsis'))
            continue
        if e.get('disagrees'):
            problems.append((key, f'recorded disagreement with the drawn label: {e["disagrees"]}'))
            continue
        labs, why = b.source(e)
        if labs is None:
            problems.append((key, why))
            continue
        if e.get('identical_to'):
            other, why = b.source(e['identical_to'])
            if other is None:
                problems.append((key, f'identical_to: {why}'))
                continue
            if other != labs:
                diff = [loc for loc in LOCS if other[loc] != labs[loc]]
                problems.append((key, f'identical_to {e["identical_to"]} differs in {diff}'))
                continue
        if e.get('strip_ellipsis'):
            labs = {loc: strip_trailing_ellipsis(v) for loc, v in labs.items()}
        rows[key] = labs
    held = {k: v for k, v in m.get('held', {}).items()}
    for key, h in held.items():
        if 'text' not in h and 'provisional' not in h:
            continue
        labs, text = None, h.get('text', {})
        if 'provisional' in h:
            labs, why = b.source(h['provisional'])
            if labs is None:
                problems.append((key, f'provisional: {why}'))
                continue
        elif 'en' not in text:
            problems.append((key, 'held text without en or a provisional source'))
            continue
        rows[key] = {loc: text.get(loc, labs[loc] if labs else text['en']) for loc in LOCS}
    return rows, problems, held


def report(rows, problems, held, ours):
    gallery, function = out_of_scope(ours)
    diff = {loc: 0 for loc in LOCS}
    for key, labs in rows.items():
        ns, _, name = key.partition(':')
        if ns not in NAMESPACES or key in held:
            continue
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
          + (' (generated)' if key in rows else '')
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
        doc = json.load(open(a.map, encoding='utf-8'))
        mapped = doc['map']
        consumers = set(doc.get('consumers', ()))
        ribbon = [k for k in rows if k in mapped and k.partition(':')[0] in NAMESPACES]
        unresolved = sorted(k for k in ribbon if k not in consumers)
        if unresolved:
            print(f'mapped keys that our label set never resolves (run verify-mac-ribbon-labels.py --write): '
                  f'{unresolved}', file=sys.stderr)
            return 1
        unverified = {loc: sum(1 for k in ribbon if loc not in mapped[k].get('verified', {})) for loc in LOCS}
        print(f'mapped rows without a drawn reading, per locale: {unverified}', file=sys.stderr)
        path = Path(a.check)
        if not path.exists() or path.read_text(encoding='utf-8') != text:
            print(f'{a.check} differs from the generated output', file=sys.stderr)
            return 1
        return 0
    Path(a.out).write_text(text, encoding='utf-8')
    return 0


if __name__ == '__main__':
    sys.exit(main())
