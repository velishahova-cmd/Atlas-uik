"""Собирает data.js для «Атласа УИК Москвы».

python build_data.py ТАБЛИЦА.xlsx Москва_Moscow.geojson data.js [new17.csv] [uik_coords.csv]

new17.csv      участки, появившиеся в источнике после снимка таблицы (разделитель ;)
uik_coords.csv координаты УИК: n;lat;lon;src;locality;street;house;inst
               src: m Мосгоризбирком (здание), o OpenStreetMap (здание),
                    r улица, p населённый пункт или микрорайон, d адреса нет (точка в центре района)
"""
import json, re, sys, collections, csv
import statistics as st
import openpyxl
from shapely.geometry import shape, mapping, Polygon, MultiPolygon
from shapely.ops import unary_union
from shapely.geometry.polygon import orient


def fix_orient(g):
    if isinstance(g, Polygon): return orient(g, -1.0)
    if isinstance(g, MultiPolygon): return MultiPolygon([orient(p, -1.0) for p in g.geoms])
    return g


XLSX = sys.argv[1]
GEO = sys.argv[2]
OUT = sys.argv[3]
NEW = sys.argv[4] if len(sys.argv) > 4 and sys.argv[4] not in ('', '-') else None
COORDS = sys.argv[5] if len(sys.argv) > 5 else None

wb = openpyxl.load_workbook(XLSX, data_only=True)

PARTIES = ['ЕР', 'ЛДПР', 'ПД', 'Зелёные', 'СР', 'Родина', 'КПРФ', 'Пенс', 'КомРос', 'НЛ']
PARTY_FULL = {'ЕР': 'Единая Россия', 'ЛДПР': 'ЛДПР', 'ПД': 'Партия прямой демократии', 'Зелёные': 'Зелёные',
              'СР': 'Справедливая Россия', 'Родина': 'Родина', 'КПРФ': 'КПРФ', 'Пенс': 'Партия пенсионеров',
              'КомРос': 'Коммунисты России', 'НЛ': 'Новые люди'}
# column indices of votes per party in sheet (0-based)
VOTE_COLS = [9, 11, 13, 15, 17, 19, 21, 23, 25, 27]

GROUPS = {'РГ 33 КПРФ': {'id': 33, 'leader': 'Кумин В. В.'}, 'РГ 34 КПРФ': {'id': 34, 'leader': 'Царихин В. А.'}}

# New Moscow aggregation: 2024 districts -> OSM administrative okrug polygons
AO_MAP = {
    'Новомосковский административный округ': ['район Коммунарка', 'район Филимонковский', 'район Щербинка'],
    'Троицкий административный округ': ['район Троицк', 'район Краснопахорский', 'район Вороново', 'район Бекасово'],
}
DISTRICT_TO_AO = {d: ao for ao, ds in AO_MAP.items() for d in ds}


def norm(s):
    s = s.lower().replace('ё', 'е')
    s = re.sub(r'\b(район|поселение|муниципальный округ|городской округ)\b', '', s)
    return re.sub(r'\W', '', s)


uiks = []
deg = []
anomalies = []
group_meta = {}
okrug_names = {}
for sheet, g in GROUPS.items():
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    mean = rows[2][2]
    group_meta[g['id']] = {'id': g['id'], 'leader': g['leader'], 'mean': round(mean, 2), 'okrugs': []}
    for r in rows[5:]:
        if r[4] == 'УИК':
            votes = [int(r[c] or 0) for c in VOTE_COLS]
            uiks.append({
                'g': g['id'], 'd': r[0], 'o': int(r[2]), 'n': int(r[3]),
                'el': int(r[5]), 'val': int(r[6]), 'inv': int(r[7]),
                'v': votes,
            })
            okrug_names[int(r[2])] = r[1]
        elif r[4] == 'ДЭГ':
            votes = [int(r[c] or 0) for c in VOTE_COLS]
            deg.append({'g': g['id'], 'o': int(r[2]), 'el': int(r[5]), 'val': int(r[6]), 'inv': int(r[7]), 'v': votes})
            okrug_names[int(r[2])] = r[1]
    # anomalies section
    start = next(i for i, r in enumerate(rows) if r[0] and str(r[0]).startswith('Статистические аномалии'))
    for r in rows[start + 2:]:
        if r[0] is None:
            break
        if not isinstance(r[0], int):
            break
        anomalies.append({'g': g['id'], 'd': r[1], 'o_name': r[2], 'n': int(r[3]), 'party': r[4], 'votes': int(r[5]),
                          'share': round(float(r[6]), 4), 'median': round(float(r[7]), 4), 'z': float(r[8]),
                          'turnout': round(float(r[9]), 3), 'why': r[10]})

# --- append UIKs that appeared in the source after the table snapshot ---
okrug_num = {v: k for k, v in okrug_names.items()}
added = 0
if NEW:
    for row in csv.DictReader(open(NEW, encoding='utf-8'), delimiter=';'):
        o = okrug_num[row['okrug']]
        g = next(gid for gid, gm in group_meta.items() if o in [u['o'] for u in uiks if u['g'] == gid])
        votes = [int(row[k]) for k in ['ER', 'LDPR', 'PD', 'Zel', 'SR', 'Rodina', 'KPRF', 'Pens', 'KomRos', 'NL']]
        uiks.append({'g': g, 'd': row['tik'], 'o': o, 'n': int(row['uik']), 'el': int(row['voters']), 'val': int(row['valid']), 'inv': int(row['invalid']), 'v': votes})
        added += 1
    uiks.sort(key=lambda u: (u['g'], u['d'].lower().replace('район ', '').replace(' район', ''), u['n']))
    # recompute group means over paper UIKs
    for gid, gm in group_meta.items():
        ks = [u['v'][6] for u in uiks if u['g'] == gid]
        gm['mean'] = round(sum(ks) / len(ks), 4)
    # recompute anomalies with the table's rule (population sd)
    anomalies = []
    for gid in group_meta:
        us = [u for u in uiks if u['g'] == gid]
        byd = collections.defaultdict(list); byo = collections.defaultdict(list)
        for u in us: byd[u['d']].append(u); byo[u['o']].append(u)
        for u in us:
            if not u['val']: continue
            turnout = (u['val'] + u['inv']) / u['el']
            grp, gname = byd[u['d']], 'район'
            if len(grp) < 8: grp, gname = byo[u['o']], 'округ'
            for pi, pname in enumerate(PARTIES):
                if pname == 'КПРФ': continue
                share = u['v'][pi] / u['val']
                shares = [x['v'][pi] / x['val'] for x in grp if x['val']]
                med = st.median(shares); mean = st.mean(shares); sd = st.pstdev(shares)
                z = (share - mean) / sd if sd > 0 else 0
                full = PARTY_FULL[pname]
                if z >= 2.5 and share - med >= 0.12 and share >= 0.40 and u['val'] >= 30:
                    anomalies.append({'g': gid, 'd': u['d'], 'o_name': okrug_names[u['o']], 'n': u['n'], 'party': full, 'votes': u['v'][pi], 'share': round(share, 4), 'median': round(med, 4), 'z': round(z, 2), 'turnout': round(turnout, 3), 'why': f"доля на {(share-med)*100:.1f} п.п. выше медианы {gname} (z={z:.1f})"})
                elif pname == 'ЕР' and share >= 0.65 and turnout >= 0.55:
                    anomalies.append({'g': gid, 'd': u['d'], 'o_name': okrug_names[u['o']], 'n': u['n'], 'party': full, 'votes': u['v'][pi], 'share': round(share, 4), 'median': round(med, 4), 'z': round(z, 2), 'turnout': round(turnout, 3), 'why': f"ЕР {share*100:.0f}% при явке {turnout*100:.0f}% на бумажном УИК"})
    anomalies.sort(key=lambda a: (a['g'], -a['z']))
    print('added', added, 'means', {k: v['mean'] for k, v in group_meta.items()}, 'anomalies', len(anomalies), file=sys.stderr)

for u in uiks:
    if u['o'] not in group_meta[u['g']]['okrugs']:
        group_meta[u['g']]['okrugs'].append(u['o'])
for gm in group_meta.values():
    gm['okrugs'].sort()

# districts
dist = collections.OrderedDict()
for u in uiks:
    if u['d'] not in dist:
        dist[u['d']] = {'name': u['d'], 'o': u['o'], 'g': u['g'], 'region': DISTRICT_TO_AO.get(u['d'], u['d'])}
districts = list(dist.values())
d_index = {d['name']: i for i, d in enumerate(districts)}
for u in uiks:
    u['di'] = d_index[u['d']]
    del u['d']
for a in anomalies:
    a['di'] = d_index[a['d']]
    del a['d']

# geometry
# The whole city is simplified as one coverage (shapely.coverage_simplify), so neighbouring
# districts keep exactly shared edges: no slivers or overlaps between them, and okrug lines,
# built as unions of the same simplified districts, sit exactly on district borders.
import shapely
geo = json.load(open(GEO, encoding='utf-8'))
cov_names = [f['properties']['district'] for f in geo['features']]
cov, acc = [], None
for f in geo['features']:
    g = shapely.make_valid(shape(f['geometry']))
    if acc is not None:
        g = shapely.make_valid(g.difference(acc))   # OSM has a few square metres of overlap; give them to the first owner
    if g.geom_type == 'GeometryCollection':
        g = unary_union([p for p in g.geoms if p.geom_type in ('Polygon', 'MultiPolygon')])
    cov.append(g)
    acc = g if acc is None else unary_union([acc, g])
cov = list(shapely.coverage_simplify(cov, 0.0003))
gn = {norm(n): fix_orient(g) for n, g in zip(cov_names, cov)}
region_names = sorted({d['region'] for d in districts})
regions = []
region_shapes = {}
for rn in region_names:
    shp = gn.get(norm(rn))
    if shp is None:
        print('NO GEOMETRY FOR', rn, file=sys.stderr)
        continue
    region_shapes[rn] = shp
    parts = AO_MAP.get(rn, [rn])
    regions.append({'type': 'Feature', 'properties': {'name': rn, 'parts': parts, 'o': dist[parts[0]]['o'], 'g': dist[parts[0]]['g']},
                    'geometry': mapping(shp)})

# okrug polygons = union of the simplified regions, no further simplification
okrugs = []
for o in sorted(okrug_names):
    shps = [region_shapes[r['properties']['name']] for r in regions if r['properties']['o'] == o]
    if not shps:
        continue
    u = shapely.make_valid(unary_union(shps))
    if u.geom_type == 'GeometryCollection':
        u = unary_union([p for p in u.geoms if p.geom_type in ('Polygon', 'MultiPolygon')])
    u = fix_orient(u)
    g = next(gm['id'] for gm in group_meta.values() if o in gm['okrugs'])
    okrugs.append({'type': 'Feature', 'properties': {'num': o, 'name': okrug_names[o], 'g': g}, 'geometry': mapping(u)})


# --- polling station locations ---
ABBR = {'УЛ.': 'ул.', 'ПРОСП.': 'просп.', 'ПР.': 'пр.', 'ПЕР.': 'пер.', 'НАБ.': 'наб.', 'БУЛЬВ.': 'бульв.',
        'ШОССЕ': 'шоссе', 'ПЛ.': 'пл.', 'ТУП.': 'туп.', 'АЛЛЕЯ': 'аллея', 'ПРОЕЗД': 'проезд', 'ПРОСЕКА': 'просека',
        'ЛИНИЯ': 'линия', 'МКР': 'мкр', 'ПОС.': 'пос.', 'ДЕР.': 'дер.'}
SETTLEMENT = re.compile(r'^(.*?) (пос\.|дер\.|с\.|г\.|дач\.пос\.|раб\.пос\.)$')


def nice_street(s):
    if not s or s.upper() != s:
        return s
    s = re.sub(r'\.(?=[А-ЯЁ])', '. ', s)
    out = []
    for t in s.split(' '):
        if t in ABBR:
            out.append(ABBR[t]); continue
        m = re.match(r'^(\d+)-(Й|Я|Е|ГО|ОЙ|ЫЙ)$', t)
        if m:
            out.append(m.group(1) + '-' + m.group(2).lower()); continue
        out.append('-'.join(p[:1] + p[1:].lower() for p in t.split('-')))
    return ' '.join(out)


def settlement(loc):
    m = SETTLEMENT.match(loc or '')
    return f'{m.group(2)} {m.group(1)}' if m else None


def address(loc, street, house):
    street = re.sub(r'\bМКР\b', 'мкр', nice_street(street))
    if (loc or '').startswith('Зеленоград'):
        return f'Зеленоград, корп. {house}' if street == loc else f'Зеленоград, {street}, {house}'
    s = settlement(loc)
    if s:
        return f'{s}, {house}' if street == loc else f'{s}, {street}, {house}'
    return f'{street}, {house}'


coords_stats = collections.Counter()
if COORDS:
    rows = {}
    for line in open(COORDS, encoding='utf-8'):
        p = line.rstrip('\n').split(';')
        if len(p) < 8 or not p[0].strip():
            continue
        rows[int(p[0])] = p[:7] + [';'.join(p[7:])]
    for u in uiks:
        r = rows.get(u['n'])
        if not r:
            coords_stats['missing'] += 1
            continue
        n, lat, lon, src, loc, street, house, inst = r
        inst = re.sub(r',?\s*тел\..*$', '', inst).strip()
        if len(inst) >= 48:
            inst = inst.rstrip(' ,.') + '…'
        if src == 'd' or not lat:
            rn = districts[u['di']]['region']
            pt = region_shapes[rn].representative_point()
            u['ll'] = [round(pt.y, 5), round(pt.x, 5)]
            u['src'] = 'd'
            u['adr'] = 'адрес помещения в открытых данных не найден'
        else:
            u['ll'] = [round(float(lat), 5), round(float(lon), 5)]
            u['src'] = src
            u['adr'] = address(loc, street, house)
            if inst:
                u['inst'] = inst
        coords_stats[u['src']] += 1
    print('coords', dict(coords_stats), file=sys.stderr)


def round_geom(obj, nd=5):
    if isinstance(obj, (list, tuple)):
        if obj and isinstance(obj[0], float):
            return [round(x, nd) for x in obj]
        return [round_geom(x, nd) for x in obj]
    return obj


for f in regions + okrugs:
    f['geometry']['coordinates'] = round_geom(f['geometry']['coordinates'])

data = {
    'meta': {
        'source': ('Протоколы федерального списка, снимок 2026-10-12, поучастковая база Выборотеки (1 459 строк по Москве: 1 443 бумажных УИК и 16 ДЭГ)' if NEW else 'Протоколы федерального списка, снимок 2026-09-28, поучастковая база Выборотеки (1 442 протокола Москвы)'),
        'source_url': 'https://www.vyboroteka.org/2026/2026-09-20/ordered/gorod-moskva?kind=gosduma&contest=party_list',
        'parties': PARTIES, 'party_full': PARTY_FULL,
        'anomaly_rule': 'Среди бумажных УИК района (если их меньше 8, то округа) доля партии имеет z ≥ 2,5 и минимум на 12 п.п. выше медианы, при доле не ниже 40% и не меньше 30 действительных бюллетеней. Отдельно отмечается ЕР от 65% при явке бумажного УИК от 55%. Это статистические выбросы, не заключение о фальсификации.',
        'ao_note': 'Семь районов Новой Москвы 2024 года (Коммунарка, Филимонковский, Щербинка; Троицк, Краснопахорский, Вороново, Бекасово) показаны двумя контурами административных округов: отдельных границ в открытых данных пока нет.',
        'coords_note': ('Адреса и координаты помещений: сервис поиска УИК Мосгоризбиркома (mosgorizbirkom.ru), снимок 2026-10-15. '
                        'Где у сервиса нет координат, адрес найден в OpenStreetMap через Nominatim (данные ODbL); '
                        'каждая такая точка сверена обратным геокодированием и попаданием в свой район.') if COORDS else '',
    },
    'groups': list(group_meta.values()),
    'okrug_names': {str(k): v for k, v in okrug_names.items()},
    'districts': districts,
    'uiks': uiks,
    'deg': deg,
    'anomalies': anomalies,
    'regions': {'type': 'FeatureCollection', 'features': regions},
    'okrugs': {'type': 'FeatureCollection', 'features': okrugs},
}
js = 'window.KPRF_DATA = ' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';\n'
open(OUT, 'w', encoding='utf-8').write(js)
print('uiks', len(uiks), 'deg', len(deg), 'anomalies', len(anomalies), 'districts', len(districts), 'regions', len(regions), 'okrugs', len(okrugs), 'bytes', len(js.encode()))
