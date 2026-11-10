#!/usr/bin/env python3
"""
build_from_powerbi.py - rebuild the Monster Brewing Field Survey Dashboard from Power BI.

Sources (Monster Brewing workspace):
  KARMA model      "MBrew KARMA Preview"                 -> surveys + buyer commitments
  Depletions model "Monster Brewing Depletions Dashboard" -> distributors (Bottler), chain flag,
                                                            case equivalents (delivery status)

Env vars (GitHub secrets):
  PBI_TENANT_ID, PBI_CLIENT_ID, PBI_CLIENT_SECRET   service principal with read access to both models
Optional:
  SURVEY_START   first survey date to include (default 2026-07-01)
  OUT_DIR        where index.html is written (default dist/)

Local dry run without Power BI:  python3 build_from_powerbi.py --fixture tests/fixture.json
"""
import argparse, base64, collections, datetime as dt, json, os, re, sys, time
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

HERE = os.path.dirname(os.path.abspath(__file__))
KARMA_MODEL = "07f05c80-03e6-49c7-8e08-0a1cc6c27e61"
DEPL_MODEL = "53e6e702-382d-4273-841c-8c01576104c1"
SURVEY_FORMS = ['MBREW B.T.B. 2026', 'MBREW B.T.B. 2026 - NY']
COMMIT_FORMS = ['Buyer Commitments', 'Buyer Commitments - UT']
SURVEY_START = os.environ.get('SURVEY_START', '2026-07-01')

# KARMA question ids (MBREW B.T.B. 2026)
Q_FU, Q_NOTES, Q_COMMENTS = 39, 2478162, 47
Q_NEWDISP, Q_BONUS, Q_WW = 2481616, 2756345, 2757696
Q_RIDING, Q_RIDING_NAME = 2754669, 2764104
PHOTO_Q = [(2759074, 'Shelf/Vault'), (2477589, 'Photo 1'), (2477590, 'Photo 2'),
           (2622395, 'New Display'), (2757698, 'W&W/POS')]
# Buyer Commitments
Q_GOT, Q_QTY, Q_LOC, Q_INSTR = 1095, 1097, 2693504, 2693505
Q_SPOKE_WITH, Q_MBC_PHONE, Q_MBC_NAME, Q_ADDL = 2617475, 2693508, 2693507, 2693428

SODA = re.compile(r'coca|cola|coke|ccbc|swire|bottling|united cc|mec direct', re.I)
AUTO_NOTE = 'Survey submitted with no brand commitments recorded — verify with the rep.'

# ---------------------------------------------------------------- helpers
def s(v):
    return '' if v is None else str(v).strip()

def tc(v):
    return s(v).title()

def karma_dt(v):          # "2026-11-01-12.24.21.000000" -> "2026-11-01 12:24"
    v = s(v)
    return f'{v[0:10]} {v[11:13]}:{v[14:16]}' if len(v) >= 16 else v

def month_windows(start, end):
    """[(first, next_first), ...] covering start..end, as ISO strings."""
    cur = dt.date.fromisoformat(start).replace(day=1)
    stop = dt.date.fromisoformat(end)
    out = []
    while cur <= stop:
        nxt = (cur.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
        out.append((cur.isoformat(), nxt.isoformat()))
        cur = nxt
    return out

def dax_list(values):
    return '{' + ','.join('"%s"' % str(v).replace('"', '') for v in values) + '}'

# ---------------------------------------------------------------- Power BI REST
class PowerBI:
    def __init__(self):
        import msal, requests
        self.requests = requests
        app = msal.ConfidentialClientApplication(
            os.environ['PBI_CLIENT_ID'],
            authority=f"https://login.microsoftonline.com/{os.environ['PBI_TENANT_ID']}",
            client_credential=os.environ['PBI_CLIENT_SECRET'])
        tok = app.acquire_token_for_client(scopes=['https://analysis.windows.net/powerbi/api/.default'])
        if 'access_token' not in tok:
            sys.exit(f"Power BI auth failed: {tok.get('error_description', tok)}")
        self.hdr = {'Authorization': 'Bearer ' + tok['access_token']}

    def query(self, dataset, dax, tries=4):
        url = f'https://api.powerbi.com/v1.0/myorg/datasets/{dataset}/executeQueries'
        body = {'queries': [{'query': dax}], 'serializerSettings': {'includeNulls': True}}
        for i in range(tries):
            r = self.requests.post(url, headers=self.hdr, json=body, timeout=300)
            if r.status_code == 429 or r.status_code >= 500:
                time.sleep(15 * (i + 1)); continue
            if r.status_code != 200:
                sys.exit(f'executeQueries {r.status_code}: {r.text[:800]}\n--- DAX ---\n{dax[:800]}')
            rows = r.json()['results'][0]['tables'][0].get('rows', [])
            # keys come back as "[alias]" - strip the brackets
            return [{k.strip('[]').split('[')[-1].rstrip(']'): v for k, v in row.items()} for row in rows]
        sys.exit('Power BI kept throttling; giving up')

# ---------------------------------------------------------------- DAX
def dax_survey(forms, d0, d1):
    return f"""EVALUATE SELECTCOLUMNS(
 FILTER('KARMA Survey', 'KARMA Survey'[SURVEYNAME] IN {dax_list(forms)}
   && 'KARMA Survey'[DATE_TAKEN] >= "{d0}" && 'KARMA Survey'[DATE_TAKEN] < "{d1}"),
 "vip", 'KARMA Survey'[VIP_OUTLET_ID], "emp", 'KARMA Survey'[TASK_TAKER_USER_NAME],
 "dt", 'KARMA Survey'[DATE_TAKEN], "q", 'KARMA Survey'[QUESTION_ID], "a", 'KARMA Survey'[Answers],
 "item", 'KARMA Survey'[ITEM_ID], "brand", 'KARMA Survey'[ITEM_DESC],
 "acct", RELATED('VIP Store'[Account Name]), "num", RELATED('VIP Store'[StoreNumber]),
 "addr", RELATED('VIP Store'[Street Address]), "city", RELATED('VIP Store'[City]),
 "st", RELATED('VIP Store'[State]), "zip", RELATED('VIP Store'[Zip]),
 "cot", RELATED('VIP Store'[VPCot]), "prem", RELATED('VIP Store'[PremiseType]),
 "dist", RELATED('VIP Store'[DistributorID]))"""

DAX_BOTTLER = """EVALUATE SELECTCOLUMNS('Bottler', "id", 'Bottler'[Distributor ID],
 "dc", 'Bottler'[Distribution Center], "parent", 'Bottler'[Bottler], "st", 'Bottler'[State])"""

def dax_chain(vips):
    return f"""EVALUATE SELECTCOLUMNS(FILTER('VIP Store', 'VIP Store'[VIPID] IN {dax_list(vips)}),
 "vip", 'VIP Store'[VIPID], "ci", 'VIP Store'[Chain/Independent])"""

def dax_caseeq(vips, d0, d1):
    return f"""EVALUATE SUMMARIZECOLUMNS('Depletion'[VIPID], 'Depletion'[DistID],
 FILTER(ALL('Depletion'[VIPID]), 'Depletion'[VIPID] IN {dax_list(vips)}),
 FILTER(ALL('Depletion'[InvoiceDate]), 'Depletion'[InvoiceDate] >= DATE({d0.replace('-', ',')})
   && 'Depletion'[InvoiceDate] <= DATE({d1.replace('-', ',')})),
 "ce", SUM('Depletion'[Case Equivs]))"""

DAX_MAXDATES = """EVALUATE ROW("survey", CALCULATE(MAX('KARMA Survey'[DATE_TAKEN]),
 'KARMA Survey'[SURVEYNAME] = "MBREW B.T.B. 2026"))"""
DAX_MAXINV = """EVALUATE ROW("inv", MAX('Depletion'[InvoiceDate]))"""

# ---------------------------------------------------------------- transforms
def group_visits(answers):
    """long rows -> {(vip, emp, dt): {'q': {qid: answer}, 'lines': {item: {...}}, 'store': {...}}}"""
    visits = {}
    for r in answers:
        key = (s(r['vip']), s(r['emp']), s(r['dt']))
        v = visits.setdefault(key, {'q': {}, 'lines': {}, 'store': r})
        q = int(r['q'])
        item = r.get('item')
        if item not in (None, 0, '0') and q in (Q_GOT, Q_QTY, Q_LOC, Q_INSTR):
            v['lines'].setdefault(str(item), {'brand': s(r.get('brand'))})[q] = s(r['a'])
        else:
            v['q'][q] = s(r['a'])
    return visits

def dist_name(code, bottler, fallback):
    code = s(code)
    if code in bottler:
        return bottler[code]['dc']
    return fallback.get(code, code)

def build_rows(visits, empmap, bottler, dist_fallback):
    rows, unknown_emps = [], collections.Counter()
    for (vip, emp, dtv), v in visits.items():
        q, st = v['q'], v['store']
        if emp in empmap:
            role = empmap[emp][0]
        else:
            role = 'Field Sales Representative'; unknown_emps[emp] += 1
        notes, com = q.get(Q_NOTES, ''), q.get(Q_COMMENTS, '')
        note = notes + (' — ' + com if notes and com else com if com else '')
        photos = [[lab, q[qid]] for qid, lab in PHOTO_Q if q.get(qid)]
        ride = tc(q.get(Q_RIDING, ''))
        if 'Mbrew' in ride and q.get(Q_RIDING_NAME):
            ride += ': ' + ' '.join(tc(q[Q_RIDING_NAME]).split())
        acct = tc(st.get('acct'))
        num = st.get('num')
        chain = st.get('chain', 'Independents')
        if num not in (None, '', 0):
            acct = f'{acct} {num}'
        d = dist_name(st.get('dist'), bottler, dist_fallback)
        d = '—' if not d or SODA.search(d) else d
        rows.append([
            emp, role, karma_dt(dtv), acct, tc(st.get('city')), s(st.get('st')).upper(),
            s(st.get('prem')).upper(), tc(st.get('cot')), chain, d,
            note, norm_yn(q.get(Q_FU)), norm_yn(q.get(Q_NEWDISP)), photos, vip,
            tc(st.get('addr')), s(st.get('zip'))[:5], ride, norm_yn(q.get(Q_BONUS)), norm_yn(q.get(Q_WW)),
        ])
    rows.sort(key=lambda r: (r[2], r[0]))
    return rows, unknown_emps

def norm_yn(v):
    v = s(v).upper()
    return 'YES' if v == 'YES' else 'NO' if v == 'NO' else s(v)

def build_buy(visits, bottler, dist_fallback, caseeq, rows):
    """Buyer Commitment visits -> BUY accounts; case_eq/status from depletions by (vip, distributor)."""
    accts = {}
    for (vip, emp, dtv), v in sorted(visits.items(), key=lambda kv: kv[0][2]):
        q, st = v['q'], v['store']
        dcode = s(st.get('dist'))
        b = bottler.get(dcode)
        dc = (b['dc'] + ', ' + b['st']) if b else dist_fallback.get(dcode, dcode)
        parent = (b['parent'] + ', ' + b['st']) if b else ''
        a = accts.setdefault(vip, dict(
            acct=tc(st.get('acct')) + (f" {st['num']}" if st.get('num') not in (None, '', 0) else ''),
            addr=tc(st.get('addr')), city=tc(st.get('city')), st=s(st.get('st')).upper(),
            zip=s(st.get('zip'))[:5], vip=vip, _dcode=dcode, blocks={}))
        blk = a['blocks'].setdefault((dc, emp), dict(
            dc=dc, parent=parent, rep=emp, case_eq=0.0, status='OPEN', dates=[], commits=[],
            contact=dict(name='', phone='', spoke=''), note=''))
        day = dtv[:10]
        if day not in blk['dates']:
            blk['dates'].append(day)
        for item, ln in v['lines'].items():
            if not ln.get('brand'):
                continue
            c = dict(brand=tc(ln['brand']), qty=int(float(ln.get(Q_QTY) or 0)), got=norm_yn(ln.get(Q_GOT)),
                     loc=s(ln.get(Q_LOC)), instr=s(ln.get(Q_INSTR)))
            if c not in blk['commits']:
                blk['commits'].append(c)
        if q.get(Q_SPOKE_WITH) or q.get(Q_MBC_PHONE) or q.get(Q_MBC_NAME):
            blk['contact'] = dict(name=tc(q.get(Q_SPOKE_WITH)), phone=s(q.get(Q_MBC_PHONE)),
                                  spoke=tc(q.get(Q_MBC_NAME)))
        if q.get(Q_ADDL):
            blk['note'] = q[Q_ADDL]
    buy = []
    for vip, a in accts.items():
        blocks = []
        for blk in a['blocks'].values():
            ce = caseeq.get((vip, a['_dcode']), 0.0)
            blk['case_eq'] = round(ce, 2)
            blk['status'] = 'FULFILLED' if ce > 0 else 'OPEN'
            blk['dates'].sort()
            if not blk['commits'] and not blk['note']:
                blk['note'] = AUTO_NOTE
            blocks.append(blk)
        a['blocks'] = blocks
        del a['_dcode']
        buy.append(a)
    buy.sort(key=lambda a: (a['st'], a['acct']))
    refresh_karma(buy, rows)
    return buy

def refresh_karma(buy, rows):
    dists_by_vip = collections.defaultdict(set)
    for r in rows:
        if r[9] != '—':
            dists_by_vip[r[14]].update(r[9].split(' · '))
    for a in buy:
        kd = sorted(dists_by_vip.get(a['vip'], []))
        a['karma_dists'] = kd
        if not kd:
            a['karma_match'] = 'none'
        else:
            dcs = {re.sub(r',\s*[A-Z]{2}$', '', b['dc']) for b in a['blocks']}
            a['karma_match'] = 'ok' if dcs & set(kd) else 'diff'
        a['open'] = any(b['status'] == 'OPEN' for b in a['blocks'])

# ---------------------------------------------------------------- page
def render(rows, buy, empmap, out_dir, subtitle=None):
    inner = open(os.path.join(HERE, 'template', 'dashboard.html'), encoding='utf-8').read()
    js = lambda v: json.dumps(v, ensure_ascii=False, separators=(',', ':'))
    inner = inner.replace('__ROWS__', js(rows)).replace('__BUY__', js(buy)).replace('__EMPMAP__', js(empmap))
    if subtitle:
        inner = re.sub(r'(<span id="siteSub">)[^<]*', r'\g<1>' + subtitle, inner)
    wrapper = open(os.path.join(HERE, 'template', 'wrapper.html'), encoding='utf-8').read()
    key = base64.b64decode(re.search(r'const MK="([^"]+)"', wrapper).group(1))
    iv = os.urandom(12)
    ct = AESGCM(key).encrypt(iv, inner.encode(), None)
    page = wrapper.replace('__PIV__', base64.b64encode(iv).decode()).replace('__PCT__', base64.b64encode(ct).decode())
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, 'index.html')
    open(path, 'w', encoding='utf-8').write(page)
    return path

# ---------------------------------------------------------------- main
def fetch_all(pbi, survey_start):
    today = dt.date.today().isoformat()
    max_survey = karma_dt(pbi.query(KARMA_MODEL, DAX_MAXDATES)[0]['survey'])[:10]
    max_inv = s(pbi.query(DEPL_MODEL, DAX_MAXINV)[0]['inv'])[:10]
    end = max(max_survey, today)
    survey, commit = [], []
    for d0, d1 in month_windows(survey_start, end):
        survey += pbi.query(KARMA_MODEL, dax_survey(SURVEY_FORMS, d0, d1))
        commit += pbi.query(KARMA_MODEL, dax_survey(COMMIT_FORMS, d0, d1))
    bottler = {s(b['id']): b for b in pbi.query(DEPL_MODEL, DAX_BOTTLER)}
    vips = sorted({s(r['vip']) for r in survey} | {s(r['vip']) for r in commit})
    chain, caseeq = {}, {}
    # 3-month window: first day of the month two months before the latest invoice, through the latest invoice
    inv = dt.date.fromisoformat(max_inv)
    w0 = (inv.replace(day=1) - dt.timedelta(days=1)).replace(day=1)
    w0 = (w0 - dt.timedelta(days=1)).replace(day=1)
    for i in range(0, len(vips), 400):
        chunk = vips[i:i + 400]
        for r in pbi.query(DEPL_MODEL, dax_chain(chunk)):
            chain[s(r['vip'])] = s(r['ci'])
        for r in pbi.query(DEPL_MODEL, dax_caseeq(chunk, w0.isoformat(), inv.isoformat())):
            caseeq[(s(r['VIPID']), s(r['DistID']))] = float(r['ce'] or 0)
    return dict(survey=survey, commit=commit, bottler=bottler, chain=chain, caseeq=caseeq,
                max_survey=max_survey, max_inv=max_inv, ce_window=[w0.isoformat(), inv.isoformat()])

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--fixture', help='JSON produced by --dump; skips Power BI')
    ap.add_argument('--dump', help='write fetched data to this JSON (for debugging)')
    ap.add_argument('--out', default=os.environ.get('OUT_DIR', 'dist'))
    ap.add_argument('--subtitle')
    a = ap.parse_args()

    if a.fixture:
        data = json.load(open(a.fixture))
        data['caseeq'] = {tuple(k.split('|')): v for k, v in data['caseeq'].items()}
    else:
        data = fetch_all(PowerBI(), SURVEY_START)
        if a.dump:
            d = dict(data); d['caseeq'] = {'|'.join(k): v for k, v in data['caseeq'].items()}
            json.dump(d, open(a.dump, 'w'))

    empmap = json.load(open(os.path.join(HERE, 'empmap.json'), encoding='utf-8'))
    dist_fallback = json.load(open(os.path.join(HERE, 'dist_codes.json'), encoding='utf-8'))
    for r in data['survey'] + data['commit']:
        r['chain'] = tc(r.get('acct')) if data['chain'].get(s(r['vip'])) == 'Chain' else 'Independents'

    visits = group_visits(data['survey'])
    rows, unknown = build_rows(visits, empmap, data['bottler'], dist_fallback)
    cvisits = group_visits(data['commit'])
    buy = build_buy(cvisits, data['bottler'], dist_fallback, data['caseeq'], rows)

    path = render(rows, buy, empmap, a.out, a.subtitle)
    blocks = sum(len(x['blocks']) for x in buy)
    openb = sum(1 for x in buy for b in x['blocks'] if b['status'] == 'OPEN')
    print(f"Surveys: {len(rows):,} visits ({rows[0][2][:10]} to {rows[-1][2][:10]}), latest in Power BI {data['max_survey']}")
    print(f"Commitments: {len(buy):,} accounts, {blocks:,} blocks, {openb:,} open; "
          f"case equivalents {data['ce_window'][0]} to {data['ce_window'][1]}")
    print(f"Wrote {path} ({os.path.getsize(path)/1e6:.1f} MB)")
    if unknown:
        msg = 'NEW EMPLOYEES not in empmap.json (shown as Field Sales Representative, no manager): ' + \
              ', '.join(f'{k} ({n})' for k, n in unknown.most_common())
        print('::warning::' + msg)
        open(os.path.join(a.out, 'new_employees.txt'), 'w').write(msg + '\n')

if __name__ == '__main__':
    main()
