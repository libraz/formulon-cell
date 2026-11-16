#!/usr/bin/env python3
"""Check every mapped Mac ribbon label against the labels the desktop app drew.

The node of a key is fixed from OUR model, never from the label text:

  - packages/formulon-cell/src/toolbar/ribbon/mac/model.ts gives, per tab, the groups and the
    commands in them. Group G, command P of tab T is the AX node 0.1.0.G.P in the tab's dump.
  - menu-manifest.ts gives the items of a menu command; the I-th item is the I-th item of the menu
    the dump opened from that command's node.
  - a command id resolves to a label key the way labels.ts does: a menu label seed of that id, else
    the label seed of the id's last segment; a tab id resolves to the label seed of that name.

A place is used in a locale only when that locale's dump has the structure the model has (same
number of chunks in the tab, same number of controls in the chunk, same number of items in the
menu). Otherwise the key is unverified in that locale; it is never searched for by text. Keys the
model cannot place (Home, galleries, group titles, controls the dump does not show) are
unverified everywhere.

Input dumps (--probes): <loc>/<n>_<tab>.json (ribbon controls), <loc>/tabs_menus.json (opened
menus), ja also tabs/main_*.json. Comparison is exact: the generated value is the cleaned
bundle string and the drawn title must equal it code point for code point.

  verify-mac-ribbon-labels.py            report only
  verify-mac-ribbon-labels.py --write    persist the result in the map:
      "verified": {locale: "<tab>: <ax path>"} on each mapped entry for the locales that agree
          (a list when the key is rendered on several controls);
      "control": "button" | "menu-item" from the AX role, and strip_ellipsis set to match;
      "consumers": the keys our label set resolves; a mapped key outside it is held "no-consumer";
      an entry that disagrees in a locale moves to `held` ("drawn-disagrees") with the drawn
          text, the bundle text, the locale and the AX path;
      a key rendered on several controls that draw different text moves to `held` ("split");
      "drawn_exceptions": {locale: text} on an entry accepts that drawn text for the locale.
"""
import argparse
import glob
import importlib.util
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
MAC = HERE.parent / 'packages/formulon-cell/src/toolbar/ribbon/mac'
DEFAULT_PROBES = '/Users/libraz/Projects/formulon/backup/probes/ribbon-review-2026-11-13/out'
TAB_IDS = ['home', 'insert', 'draw', 'pageLayout', 'formulas', 'data', 'review', 'view', 'automate']


def load_generator():
    spec = importlib.util.spec_from_file_location('gen_mac_ribbon_labels', HERE / 'gen-mac-ribbon-labels.py')
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def read_json(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


# ---------------------------------------------------------------- TypeScript reading (read-only)


def match_close(s, i, open_c, close_c):
    """Index of the bracket closing the one at s[i], skipping strings and line comments."""
    depth, n, j = 0, len(s), i
    while j < n:
        c = s[j]
        if c in '\'"`':
            j += 1
            while j < n and s[j] != c:
                j += 2 if s[j] == '\\' else 1
        elif s.startswith('//', j):
            while j < n and s[j] != '\n':
                j += 1
        elif c == open_c:
            depth += 1
        elif c == close_c:
            depth -= 1
            if depth == 0:
                return j
        j += 1
    raise ValueError('unbalanced brackets')


def split_top(s):
    """Split a bracket body on its top-level commas."""
    out, depth, cur, j, n = [], 0, [], 0, len(s)
    while j < n:
        c = s[j]
        if c in '\'"`':
            k = j + 1
            while k < n and s[k] != c:
                k += 2 if s[k] == '\\' else 1
            cur.append(s[j:k + 1])
            j = k + 1
            continue
        if s.startswith('//', j):
            k = s.find('\n', j)
            j = n if k < 0 else k
            continue
        if c in '([{':
            depth += 1
        elif c in ')]}':
            depth -= 1
        if c == ',' and depth == 0:
            out.append(''.join(cur).strip())
            cur = []
        else:
            cur.append(c)
        j += 1
    tail = ''.join(cur).strip()
    if tail:
        out.append(tail)
    return out


def seed_keys(src, name):
    """Keys of `const <name>: ... = { key: {...}, ... }`."""
    m = re.search(rf'const {name}\b[^=]*=\s*\{{', src)
    start = m.end() - 1
    body = src[start + 1:match_close(src, start, '{', '}')]
    return [re.sub(r"^'|'$", '', e.split(':', 1)[0].strip()) for e in split_top(body)]


def read_model():
    """tab id -> [{'title', 'cmds': [command id]}], in render order."""
    s = (MAC / 'model.ts').read_text(encoding='utf-8')
    family = re.findall(r"'(\w+)'", re.search(r'const FORMULA_FAMILY_KEYS = \[(.*?)\]', s, re.S).group(1))
    tabs = {}
    for m in re.finditer(r"const \w+Tab = \(\): RibbonTabModel =>\s*macTab\(\s*'(\w+)',", s):
        op = s.index('(', s.rindex('macTab', 0, m.end()))
        args = split_top(s[op + 1:match_close(s, op, '(', ')')])
        groups = []
        for g in split_top(args[2][1:-1]):
            if not g.startswith('group('):
                raise ValueError(f'model.ts: unexpected group expression {g[:40]!r}')
            gargs = split_top(g[len('group('):-1])
            title = re.search(r"en: '([^']*)'", gargs[0])
            cmds = []
            for c in split_top(gargs[1][1:-1]):
                if c.startswith('...FORMULA_FAMILY_KEYS'):
                    cmds += [f'mac.formulas.{k}' for k in family]
                    continue
                mm = re.match(r"\w+\(\s*(?:'([^']+)'|`([^`]+)`)", c)
                if not mm:
                    raise ValueError(f'model.ts: unexpected command expression {c[:50]!r}')
                cmds.append(mm.group(1))
            groups.append({'title': title.group(1) if title else None, 'cmds': cmds})
        tabs[m.group(1)] = groups
    return tabs


def read_menus():
    """menu command id -> [{'id', 'own'}] for the menus written out literally in menu-manifest.ts."""
    s = (MAC / 'menu-manifest.ts').read_text(encoding='utf-8')
    m = re.search(r'export const MAC_RIBBON_MENU_ITEMS[^=]*=\s*\{', s)
    start = m.end() - 1
    out = {}
    for entry in split_top(s[start + 1:match_close(s, start, '{', '}')]):
        key, _, expr = entry.partition(':')
        key = key.strip().strip("'")
        expr = expr.strip()
        wrapped = expr.startswith('menu([')
        if not wrapped and not expr.startswith('['):
            continue
        arr = expr[expr.index('['):]
        items = []
        for it in split_top(arr[1:match_close(arr, 0, '[', ']')]):
            mm = re.match(r"leaf\(\s*'([^']+)'", it) or re.search(r"\bid:\s*'([^']+)'", it)
            if not mm:
                items = None
                break
            items.append({'id': mm.group(1), 'own': (not wrapped) or bool(re.search(r'\bown:\s*true', it))})
        if items is not None:
            out[key] = items
    return out


class Surface:
    """What our label set resolves, and where our model renders each key."""

    def __init__(self):
        labels = (MAC / 'labels.ts').read_text(encoding='utf-8')
        self.label_keys = set(seed_keys(labels, 'labelSeeds'))
        self.menu_keys = set(seed_keys(labels, 'menuLabelSeeds'))
        self.group_keys = set(seed_keys(labels, 'groupSeeds'))
        self.model = read_model()
        self.menus = read_menus()
        text = ''.join((MAC / f).read_text(encoding='utf-8')
                       for f in ('model.ts', 'menu-manifest.ts', 'menus.ts', 'shape-gallery.ts', 'object-tabs.ts',
                                 'state-projection.ts', 'dispatch.ts', 'actions.ts', 'automation-gallery.ts', 'labels.ts'))
        self.ids = set(re.findall(r"'((?:mac\.)[\w.]+)'", text))
        self.ids |= {c for gs in self.model.values() for g in gs for c in g['cmds']}
        self.ids |= {i['id'] for items in self.menus.values() for i in items}
        self.tail = {i.rsplit('.', 1)[-1] for i in self.ids}
        direct = set(re.findall(r"\bt\('(\w+)'\)", text)) | set(re.findall(r"\blabels(?:\.|\[')(\w+)", text))
        direct |= set(re.findall(r"\bleaf\(\s*'[^']+',\s*'(\w+)'", text))
        self.direct = direct
        model_src = (MAC / 'model.ts').read_text(encoding='utf-8')
        self.group_titles = set(re.findall(r"group\(\s*\{[^}]*\ben: '([^']+)'", model_src))
        # The shared Home tab's groups are keyed by position; chart tab groups name their key last.
        home = re.search(r"MAC_HOME_GROUP_KEYS = \[(.*?)\]", model_src, re.S)
        self.group_titles |= set(re.findall(r"'([^']+)'", home.group(1))) if home else set()
        objects = (MAC / 'object-tabs.ts').read_text(encoding='utf-8')
        self.group_titles |= set(re.findall(r"\],\s*'([^']+)',?\s*\)", objects))

    def key_of(self, cid):
        """The label key a command id renders with (labels.ts labelFor), or None."""
        if cid in self.menu_keys:
            return f'menu:{cid}'
        last = cid.rsplit('.', 1)[-1]
        return f'labels:{last}' if last in self.label_keys else None

    def resolves(self, key):
        ns, _, name = key.partition(':')
        if ns == 'menu':
            return name in self.menu_keys and name in self.ids
        if ns == 'labels':
            return name in self.label_keys and (name in self.tail or name in self.direct or name in TAB_IDS)
        return name in self.group_keys and name in self.group_titles


# ---------------------------------------------------------------- the drawn dumps


class Drawn:
    """One locale's dump: titles and roles by address, and the structure around them.

    Addresses are (tab index, menu button path, item path); a ribbon control has no menu path.
    The per-tab and opened-menu dumps keep every code point; labels.json is the same set with
    zero-width spaces dropped, so it is not read here.
    """

    def __init__(self, probes, loc):
        base = Path(probes) / loc
        menus = read_json(base / 'tabs_menus.json')
        self.tabs = list(menus)
        tag_idx = {v.get('tag'): i for i, v in enumerate(menus.values())}
        idx = {name: i for i, name in enumerate(self.tabs)}
        self.at, self.role = {}, {}
        self.chunks = {}  # tab -> chunk -> [child paths]
        self.menu = {}    # (tab, button path) -> [item paths]

        files = glob.glob(str(base / '*_*.json'))
        if loc == 'ja':
            files += glob.glob(str(Path(probes) / 'tabs' / 'main_*.json'))
        for f in files:
            tag = Path(f).stem
            tag = tag.split('main_', 1)[1] if tag.startswith('main_') else tag
            if tag not in tag_idx:
                continue
            tab = tag_idx[tag]
            for r in read_json(f).get('rows', []):
                p = r['path'].split('.')
                if p[:3] != ['0', '1', '0']:
                    if len(p) == 3 and p[:2] == ['0', '1'] and p[2].isdigit():
                        self._put((tab, '', r['path']), r.get('title'), r.get('role'))
                    continue
                if len(p) == 5:
                    lst = self.chunks.setdefault(tab, {}).setdefault(int(p[3]), [])
                    if int(p[4]) not in lst:
                        lst.append(int(p[4]))
                    self._put((tab, '', r['path']), r.get('title'), r.get('role'))
        for name, v in menus.items():
            for mn in v.get('menus', []):
                lst = []
                for it in mn.get('menu', []):
                    if it.get('role') == 'AXMenuItem' and it.get('title'):  # an untitled item is a separator
                        self._put((idx[name], mn['button']['path'], it['path']), it.get('title'), 'AXMenuItem')
                        lst.append(it['path'])
                self.menu[(idx[name], mn['button']['path'])] = sorted(lst, key=lambda x: [int(n) for n in x.split('.')])
        for tab in self.chunks:
            for g in self.chunks[tab]:
                self.chunks[tab][g].sort()

    def _put(self, addr, title, role):
        self.role.setdefault(addr, role)
        if title:
            self.at.setdefault(addr, set()).add(title)

    def tab(self, i):
        return self.tabs[i]

    def control(self, tab, g, ncmds, nchunks):
        """Path of the control list of chunk g when the dump has the model's structure, else None."""
        ch = self.chunks.get(tab, {})
        if len(ch) != nchunks or len(ch.get(g, [])) != ncmds:
            return None
        return [f'0.1.0.{g}.{p}' for p in ch[g]]


def ax_text(d, addr):
    return f'{d.tab(addr[0])}: ' + (f'{addr[1]} > {addr[2]}' if addr[1] else addr[2])


# ---------------------------------------------------------------- places


def places_of(surface):
    """key -> [place]; a place is {'tab', 'g', 'p', 'menu': item index or None, 'nmenu', 'ncmds', 'nchunks', 'cmd'}."""
    out = {}

    def add(key, place):
        if key:
            out.setdefault(key, []).append(place)

    for i, tid in enumerate(TAB_IDS):
        if tid in surface.label_keys:
            add(f'labels:{tid}', {'kind': 'strip', 'tab': i, 'path': f'0.1.{2 + i}', 'cmd': f'tab {tid}'})
    for tid, groups in surface.model.items():
        tab = TAB_IDS.index(tid)
        for g, grp in enumerate(groups):
            for p, cid in enumerate(grp['cmds']):
                base = {'tab': tab, 'g': g, 'p': p, 'ncmds': len(grp['cmds']), 'nchunks': len(groups), 'cmd': cid}
                add(surface.key_of(cid), {**base, 'kind': 'control'})
                items = surface.menus.get(cid)
                if items:
                    for k, it in enumerate(items):
                        if not it['own']:
                            add(surface.key_of(it['id']),
                                {**base, 'kind': 'item', 'item': k, 'nitems': len(items), 'cmd': it['id'], 'menu_of': cid})
    return out


def resolve(drawn, place):
    """Address of a place in one locale's dump, or None when the dump has a different structure."""
    if place['kind'] == 'strip':
        return (place['tab'], '', place['path'])
    paths = drawn.control(place['tab'], place['g'], place['ncmds'], place['nchunks'])
    if paths is None:
        return None
    btn = paths[place['p']]
    if place['kind'] == 'control':
        return (place['tab'], '', btn)
    items = drawn.menu.get((place['tab'], btn), [])
    if len(items) != place['nitems']:
        return None
    return (place['tab'], btn, items[place['item']])


# ---------------------------------------------------------------- verification


def verify(gen, app, mapfile, probes):
    m = read_json(mapfile)
    surface = Surface()
    b = gen.Bundle(app)
    rows, problems, _ = gen.build(b, None, mapfile)
    if problems:
        raise SystemExit(f'generator problems: {problems}')
    locs = list(gen.LOCS)
    drawn = {loc: Drawn(probes, loc) for loc in locs}
    places = places_of(surface)
    result = {}
    for key, labs in rows.items():
        # Held rows and the non-ribbon namespaces have no drawn control to compare with.
        if key not in m['map'] or key.partition(':')[0] not in gen.NAMESPACES:
            continue
        e = m['map'][key]
        exc = e.get('drawn_exceptions', {})
        ps = places.get(key, [])
        per_place = []  # [(place, {loc: ('ok'|'bad'|'none', addr, drawn text)})]
        for pl in ps:
            res = {}
            for loc in locs:
                addr = resolve(drawn[loc], pl)
                titles = drawn[loc].at.get(addr) if addr else None
                if not titles:
                    res[loc] = ('none', addr, None)
                elif exc.get(loc, labs[loc]) in titles:
                    res[loc] = ('ok', addr, exc.get(loc, labs[loc]))
                else:
                    res[loc] = ('bad', addr, sorted(titles)[0])
            per_place.append((pl, res))
        agree, disagree, missing = {}, {}, []
        for loc in locs:
            seen = [(pl, r[loc]) for pl, r in per_place if r[loc][0] != 'none']
            if not seen:
                missing.append(loc)
                continue
            oks = [ax_text(drawn[loc], a) for _, (s, a, _) in seen if s == 'ok']
            bads = [(d, ax_text(drawn[loc], a)) for _, (s, a, d) in seen if s == 'bad']
            if bads and not oks:
                disagree[loc] = bads[0]
            elif oks and not bads:
                agree[loc] = oks[0] if len(oks) == 1 else oks
            else:
                disagree[loc] = bads[0]
        # several controls drawing different text for one key
        split = []
        if len(per_place) > 1:
            for loc in locs:
                texts = {}
                for pl, r in per_place:
                    if r[loc][0] != 'none':
                        texts.setdefault(r[loc][2], []).append(f'{pl["cmd"]} @ {ax_text(drawn[loc], r[loc][1])}')
                if len(texts) > 1:
                    split.append((loc, texts))
        roles = {drawn['en'].role.get(r['en'][1]) for _, r in per_place if r['en'][0] != 'none'}
        control = None
        if key.startswith('menu:'):
            control = 'menu-item'
        elif key.startswith('labels:') and roles:
            control = 'menu-item' if roles == {'AXMenuItem'} else 'button'
        result[key] = {'places': len(ps), 'agree': agree, 'disagree': disagree, 'missing': missing,
                       'split': split, 'labs': labs, 'control': control, 'mixed_controls': len(roles) > 1}
    return m, result, surface


def save(path, m):
    m['map'] = dict(sorted(m['map'].items()))
    m['held'] = dict(sorted(m['held'].items()))
    Path(path).write_text(json.dumps(m, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--app', default='/Applications/Microsoft Excel.app')
    p.add_argument('--probes', default=DEFAULT_PROBES)
    p.add_argument('--map', default=str(HERE / 'mac-ribbon-label-tcids.json'))
    p.add_argument('--write', action='store_true')
    p.add_argument('--json', metavar='PATH', help='dump the raw result')
    a = p.parse_args()
    gen = load_generator()
    m, res, surface = verify(gen, a.app, a.map, a.probes)
    locs = list(gen.LOCS)
    nocons = sorted(k for k in res if not surface.resolves(k))
    bad = {k: r for k, r in res.items() if k not in nocons and (r['disagree'] or r['split'])}
    unv = {loc: sum(1 for k, r in res.items() if loc in r['missing']) for loc in locs}
    print(f'mapped {len(res)}; no consumer {len(nocons)}; disagreeing/split rows {len(bad)}; unverified per locale {unv}',
          file=sys.stderr)
    for k in nocons:
        print(f'NO-CONSUMER {k}', file=sys.stderr)
    for k, r in sorted(bad.items()):
        for loc, (d, path) in r['disagree'].items():
            print(f'DISAGREE {k} [{loc}] {path}: drawn {d!r} bundle {r["labs"][loc]!r}', file=sys.stderr)
        for loc, texts in r['split']:
            print(f'SPLIT {k} [{loc}] {texts}', file=sys.stderr)
    for k, r in res.items():
        if r['mixed_controls']:
            print(f'MIXED-CONTROLS {k}', file=sys.stderr)
    if a.json:
        Path(a.json).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding='utf-8')
    if not a.write:
        return 0
    consumers = sorted(k for k in res if surface.resolves(k))
    # Pass 1: control and the ellipsis flag it implies; that changes the generated text, so the
    # drawn comparison is repeated before the result is persisted.
    for key, r in res.items():
        e = m['map'][key]
        if r['control']:
            e['control'] = r['control']
        if key.startswith('labels:') and r['control'] == 'button':
            e['strip_ellipsis'] = True
        elif r['control'] == 'menu-item':
            e.pop('strip_ellipsis', None)
    save(a.map, m)
    m, res, surface = verify(gen, a.app, a.map, a.probes)
    for key, r in res.items():
        e = m['map'].get(key)
        e.pop('verified', None)
        reason = None
        status = None
        if not surface.resolves(key):
            status, reason = 'no-consumer', 'nothing in the label set resolves this key'
        elif r['split']:
            status = 'split'
            reason = 'one key drawn differently on several controls (it needs one key per control): ' + '; '.join(
                f'{loc}: ' + ' | '.join(f'{t!r} at {w}' for t, w in texts.items()) for loc, texts in r['split'])
        elif r['disagree']:
            status = 'drawn-disagrees'
            origin = (f'{e["file"]} key {e["key"]}' if 'file' in e else f'{e["source"]} of TCID {e["tcid"]}')
            reason = f'{origin} disagrees with the drawn label: ' + '; '.join(
                f'{loc} at AX {path}: drawn {d!r}, bundle {r["labs"][loc]!r}' for loc, (d, path) in r['disagree'].items())
        if status:
            del m['map'][key]
            m['held'][key] = {'status': status, 'tcid': e.get('tcid'), 'name': e.get('name'), 'reason': reason}
        elif r['agree']:
            e['verified'] = r['agree']
    m['consumers'] = consumers
    save(a.map, m)
    return 0


if __name__ == '__main__':
    sys.exit(main())
