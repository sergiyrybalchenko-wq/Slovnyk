#!/usr/bin/env python3
"""Ставить посилання на джерела в текст визначень редакції Сергія.

Вхід:
  redaktsiia/slovnyk-redaktsiia-serhiia-2026-11-20.docx — редакція Сергія (база, вручну не правиться)
  dani/posylannia.tsv, dani/novi.tsv, dani/rishennia.tsv, dani/statti.tsv, dani/dzherela-novi.tsv
Вихід:
  build/slovnyk-nova-redaktsiia.docx
  build/posylannia-zvit.tsv — яке посилання біля якої статті стоїть (для перевірки)

Посилання — номер джерела з переліку «Використані джерела» у квадратних дужках перед
кінцевою крапкою визначення: «… текст [36].» Звичайний текст, без гіперпосилань
(Сергій прибрав гіперпосилання з книжки). На одну статтю — одне посилання.

Додатково виправляє технічні дефекти бази (див. FIX_* нижче)."""
import copy, csv, re, sys, os
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(__file__))
from dzherela import canon, d_number  # noqa: E402

BASE = 'redaktsiia/slovnyk-redaktsiia-serhiia-2026-11-20.docx'
OUT = 'build/slovnyk-nova-redaktsiia.docx'
# Авторські джерела, що є в переліку: назва в posylannia.tsv -> номер у переліку.
AUTHOR_SRC = {'NIST AI 100-2': 35, 'USAID Digital Strategy 2020–2024 (Digital Ecosystem Framework)': 'USAID'}
# Сліди виносок Вікіпедії у вихідному тексті (номери чужих приміток, що збивали б нумерацію джерел).
WIKI_NOTES = [('мережі.[38]', 'мережі.'), ('час[261])', 'час)'), ('класифікатора.[47]', 'класифікатора.')]

NOTE = ('Число у квадратних дужках наприкінці статті, наприклад [36] (словник термінів у сфері '
        'штучного інтелекту Міністерства цифрової трансформації України), — номер у переліку '
        '«Використані джерела» праці, на основі якої підготовлено статтю. Статті без номера подано '
        'за авторським текстом словника або з редакційними уточненнями.')

# Технічні спотворення в описах джерел 1–35 (наслідок розпізнавання PDF): лише URL і літери.
SRC_FIXES = [
    ('cs-230/ cheatsheet-recurrent-neuralnetworks/', 'cs-230/cheatsheet-recurrent-neural-networks/'),
    ('https://www. coe. int/en/web/artificial-intelligence/the', 'https://www.coe.int/en/web/artificial-intelligence/the-framework-convention-on-artificial-intelligence'),
    ('fortrustworthy', 'for trustworthy'), ('https://digitalstrategy.', 'https://digital-strategy.'),
    ('https:// digital-strategy', 'https://digital-strategy'), ('assessment-listtrustworthy', 'assessment-list-trustworthy'),
    ('lbdl.pdffAccessed 10 Jul. 2023]', 'lbdl.pdf'), ('13.oldsworth', '13.Holdsworth'),
    ('ai_govemance', 'ai_governance'), ('https:// www.iso.org', 'https://www.iso.org'),
    ('iso.Org/obp/ui/#iso:std:iso-iec:22989:dis:ed-l:vl:en', 'iso.org/obp/ui/#iso:std:iso-iec:22989:dis:ed-1:v1:en'),
    ('J i, J.', 'Ji, J.'), ('doi:https:// doi.org/10.6028/nist.ai.600-l', 'doi:https://doi.org/10.6028/NIST.AI.600-1'),
    ('legalinstmments', 'legalinstruments'), ('hbr.Org/2021/l 1/managing-ai-decision-makingtools', 'hbr.org/2021/11/managing-ai-decision-making-tools'),
    ('unesdoc.unesco.Org', 'unesdoc.unesco.org'),
]

rd = lambda p: list(csv.DictReader(open(p, encoding='utf-8'), delimiter='\t'))
norm = lambda t: re.sub(r'\s+', ' ', t.replace('\xa0', ' ')).strip().lower()
W = lambda tag: qn('w:' + tag)


def ptext(el):
    return ''.join(t.text or '' for t in el.iter(W('t')))


def replace_in_runs(p, old, new):
    """Заміна тексту в абзаці, навіть якщо він розбитий на кілька w:t; повертає кількість замін."""
    n = 0
    while True:
        ts = list(p._element.iter(W('t')))
        full = ''.join(t.text or '' for t in ts)
        k = full.find(old)
        if k < 0:
            return n
        n += 1
        pos, end, first = 0, k + len(old), True
        for t in ts:
            a, b = pos, pos + len(t.text or '')
            pos = b
            if b <= k or a >= end:
                continue
            s, e = max(k, a) - a, min(end, b) - a
            t.text = t.text[:s] + (new if first else '') + t.text[e:]
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            first = False


def main():
    doc = docx.Document(BASE)
    body = doc.element.body
    P = doc.paragraphs
    log = []

    def idx(pred, start=0):
        return next(i for i, p in enumerate(P) if i >= start and pred(norm(p.text)))

    st = idx(lambda t: 'основна частина' in t, 100)
    ab = idx(lambda t: t.startswith('розшифровка складних'), st)
    ix = idx(lambda t: t.startswith('абетковий показчик'), ab)
    sx = idx(lambda t: t.startswith('використані джерела'), ix)

    # FIX 1. Залишки старого тексту (гіперпосилання вихідного файлу), приклеєні перед терміном
    # у 9 відредагованих статтях: «AdapterАдаптер – Adapter – …».
    dec = {r['id']: r for r in rd('dani/rishennia.tsv')}
    firsts = [norm(r['новий_текст'].split('\\n')[0])[:40] for r in dec.values() if r['дія'] == 'виправити']
    for p in P[st:ab]:
        t = norm(p.text)
        for f in firsts:
            k = t.find(f)
            if k > 0:
                while not norm(p.text).startswith(f):
                    ch = [c for c in p._element if c.tag != W('pPr')]
                    ch[0].getparent().remove(ch[0])
                log.append(('виправлено', 'приклеєний залишок перед терміном', p.text[:40]))
                break

    # FIX 3. Сліди виносок Вікіпедії.
    for p in P[st:ix]:
        for old, new in WIKI_NOTES:
            if old in p.text and replace_in_runs(p, old, new):
                log.append(('виправлено', f'прибрано виноску {old}', p.text[:40]))

    # FIX 3a. Зайвий знак у кінці статті «Пісочниця».
    for p in P[st:ab]:
        if norm(p.text).startswith('пісочниця') and p.text.rstrip().endswith('системи/'):
            replace_in_runs(p, 'системи/', 'системи.')
            log.append(('виправлено', 'зайвий «/» у кінці', p.text[:40]))

    # FIX 4. Статті на латинську літеру (Anthropic, KV-кеш) стояли після «Я» — переносимо
    # у розділи А і К, як у покажчику Сергія (X-ризик автор уже тримає в Х).
    P = doc.paragraphs
    letters = {norm(p.text).upper(): p for p in P[st:ab] if re.fullmatch(r'[А-ЯІЇЄҐ]', p.text.strip())}
    order = list(letters)
    for prefix, L in (('anthropic –', 'А'), ('kv-кеш;', 'К')):
        p = next(p for p in P[st:ab] if norm(p.text).startswith(prefix))
        nxt = letters[order[order.index(L) + 1]]._element
        # останній непорожній абзац перед заголовком наступної літери
        prev = nxt.getprevious()
        while prev is not None and not ptext(prev).strip():
            prev = prev.getprevious()
        prev.addnext(p._element)
        log.append(('перенесено', f'у розділ {L}', p.text[:40]))

    # FIX 5. Покажчик: «Кеш ключів і значень» — на своє абеткове місце в «К».
    P = doc.paragraphs
    ix = next(i for i, p in enumerate(P) if norm(p.text).startswith('абетковий показчик'))
    sx = next(i for i, p in enumerate(P) if norm(p.text).startswith('використані джерела'))
    sys.path.insert(0, os.path.dirname(__file__))
    from zbirka import sort_key
    kesh = next((p for p in P[ix:sx] if norm(p.text) == 'кеш ключів і значень'), None)
    if kesh is not None:
        kp = [p for p in P[ix:sx] if p.text.strip()[:1] == 'К' and len(p.text.strip()) > 1 and p is not kesh]
        after = None
        for p in kp:
            if sort_key(p.text.strip()) < sort_key(kesh.text.strip()):
                after = p
        after._element.addnext(kesh._element)
        log.append(('перенесено', 'покажчик: Кеш ключів і значень', ''))

    # Перелік джерел: нові джерела (від № 36) переписуються з dani/dzherela-novi.tsv,
    # щоб номери в тексті й переліку завжди збігалися.
    P = doc.paragraphs
    sx = next(i for i, p in enumerate(P) if norm(p.text).startswith('використані джерела'))
    olds = [p for p in P[sx + 1:] if re.match(r'\s*(\d+)\.', p.text) and int(re.match(r'\s*(\d+)', p.text).group(1)) >= 36]
    tmpl = olds[0]._element
    anchor = olds[0]._element.getprevious()
    for p in olds:
        p._element.getparent().remove(p._element)
    for s in rd('dani/dzherela-novi.tsv'):
        ne = copy.deepcopy(tmpl); anchor.addnext(ne); anchor = ne
        q = Paragraph(ne, doc)
        q.runs[0].text = f"{s['номер']}.{s['текст']}"
        for r in q.runs[1:]:
            r._element.getparent().remove(r._element)
    P = doc.paragraphs
    sx = next(i for i, p in enumerate(P) if norm(p.text).startswith('використані джерела'))
    for p in P[sx + 1:]:
        for old, new in SRC_FIXES:
            if replace_in_runs(p, old, new):
                log.append(('виправлено', f'джерело: {old[:30]}', p.text[:30]))
    src_par = {}
    for p in P[sx + 1:]:
        m = re.match(r'\s*(\d+)\.', p.text)
        if m:
            src_par.setdefault(int(m.group(1)), p)
    num = {r['текст'].split('. [online]')[0].rstrip('.'): int(r['номер']) for r in rd('dani/dzherela-novi.tsv')}
    usaid = next((n for t, n in num.items() if t.startswith('USAID')), None)

    # Статті основної частини: заголовний абзац — той, що починається з терміна покажчика і тире.
    P = doc.paragraphs
    st = next(i for i, p in enumerate(P) if 'основна частина' in norm(p.text) and i > 100)
    ab = next(i for i, p in enumerate(P) if norm(p.text).startswith('розшифровка складних'))
    ix = next(i for i, p in enumerate(P) if norm(p.text).startswith('абетковий показчик'))
    sx = next(i for i, p in enumerate(P) if norm(p.text).startswith('використані джерела'))
    terms = {norm(p.text) for p in P[ix + 1:sx] if len(p.text.strip()) > 1}
    terms |= {norm(r['термін_uk']) for r in rd('dani/novi.tsv')}  # покажчик може ділити термін на два рядки
    cyr = lambda t: t.replace('top-p', 'топ-п').replace('top-k', 'топ-к')
    terms |= {cyr(t) for t in terms}
    terms = sorted(terms, key=len, reverse=True)
    heads = {}
    for i in range(st + 1, ab):
        t = norm(P[i].text)
        for term in terms:
            if t.startswith(term) and re.match(r'\s*[-–—]', t[len(term):]):
                heads.setdefault(term, i)
                break
    head_idx = set(heads.values())

    S = {r['id']: r for r in rd('dani/statti.tsv')}
    novi = {r['термін_uk'].strip(): r for r in rd('dani/novi.tsv')}
    plan = []
    for r in rd('dani/posylannia.tsv'):
        k = r['id_або_термін']
        if k.startswith('нова: '):
            term = k[6:].strip()
            n = novi[term]
            t, _ = canon(n['джерело'], n['url'])
            if t:
                no = num.get(t)
            else:
                no = d_number(n['джерело'])
        else:
            sid = k.split()[0]
            d = dec.get(sid)
            term = re.split(r'\s+[-–—]\s*', d['новий_текст'])[0] if d and d['дія'] == 'виправити' else S[sid]['термін_uk']
            src = r['джерело']
            if 'Мінцифри' in src:
                no = 36
            elif 'Д21' in r['перевірено']:
                no = 21
            else:
                no = AUTHOR_SRC.get(src)
                if no == 'USAID':
                    no = usaid
        i = heads.get(norm(term)) or heads.get(cyr(norm(term)))
        plan.append((k, term, i, no, r['джерело']))

    report = []
    for k, term, i, no, src in plan:
        if i is None or no is None or no not in src_par:
            report.append([k, term, '', '', 'не поставлено: ' + ('немає статті' if i is None else 'джерела немає в переліку'), src])
            continue
        # Цільовий абзац: заголовний; якщо він закінчується двокрапкою — останній абзац статті.
        tgt = i
        if P[i].text.rstrip().endswith(':'):
            j = i + 1
            while j < ab and j not in head_idx and len(P[j].text.strip()) > 1:
                tgt = j
                j += 1
        add_cite(P[tgt], no)
        report.append([k, term, no, P[tgt].text[-60:], 'поставлено', src])

    # Кирилиця в українських назвах вибірок, як у редакції Сергія: «вибірка топ-к», «вибірка топ-п».
    for p in doc.paragraphs:
        t = norm(p.text)
        if t.startswith(('вибірка top-k', 'ядерна вибірка; вибірка top-p')):
            n = replace_in_runs(p, 'Вибірка top-k', 'Вибірка топ-к') + replace_in_runs(p, 'вибірка top-p', 'вибірка топ-п')
            if n:
                log.append(('виправлено', 'top-k/top-p -> топ-к/топ-п в українському терміні', p.text[:40]))

    # Гіперпосилання (успадковані з вихідного файлу) прибираємо, текст лишаємо.
    nh = 0
    for h in list(doc.element.body.iter(W('hyperlink'))):
        for ch in list(h):
            for rs in ch.iter(W('rStyle')):
                rs.getparent().remove(rs)
            h.addprevious(ch)
        h.getparent().remove(h)
        nh += 1
    log.append(('виправлено', f'прибрано гіперпосилань: {nh}', ''))

    # Примітка про посилання — окремим абзацом у кінці передмови.
    P = doc.paragraphs
    pref = next(p for p in P if norm(p.text).startswith('з огляду на динамічний розвиток'))
    ne = copy.deepcopy(pref._element); pref._element.addnext(ne)
    q = Paragraph(ne, doc)
    q.runs[0].text = NOTE
    for r in q.runs[1:]:
        r._element.getparent().remove(r._element)

    os.makedirs('build', exist_ok=True)
    doc.save(OUT)
    with open('build/posylannia-zvit.tsv', 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f, delimiter='\t', lineterminator='\n')
        w.writerow(['стаття', 'термін', 'номер', 'кінець абзацу', 'стан', 'джерело'])
        w.writerows(report)
        for l in log:
            w.writerow(['', l[2], '', '', l[0] + ': ' + l[1], ''])
    done = sum(1 for r in report if r[4] == 'поставлено')
    print('посилань поставлено:', done, 'не поставлено:', len(report) - done, 'технічних виправлень:', len(log))


BM_ID = [1000]


def add_bookmark(p, n):
    name = f'dzherelo_{n}'
    s = OxmlElement('w:bookmarkStart'); s.set(W('id'), str(BM_ID[0])); s.set(W('name'), name)
    e = OxmlElement('w:bookmarkEnd'); e.set(W('id'), str(BM_ID[0]))
    BM_ID[0] += 1
    ppr = p._element.find(W('pPr'))
    if ppr is not None:
        ppr.addnext(s)
    else:
        p._element.insert(0, s)
    p._element.append(e)


def add_cite(p, n):
    """« [n]» перед кінцевою крапкою (або в кінці, якщо крапки немає); [n] — посилання на закладку."""
    runs = [r for r in p._element.iter(W('r')) if ''.join(t.text or '' for t in r.iter(W('t'))).strip()]
    last = runs[-1]
    tlist = list(last.iter(W('t')))
    t = tlist[-1]
    txt = (t.text or '').rstrip()
    tail = ''
    if txt.endswith('.') and not txt.endswith('...'):
        txt, tail = txt[:-1], '.'
    elif txt and txt[-1] not in '.!?…:;':
        tail = '.'  # уніфікація: «… текст [N].»
    t.text = txt
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    rpr = last.find(W('rPr'))

    def mk(text):
        r = OxmlElement('w:r')
        if rpr is not None:
            rp = copy.deepcopy(rpr)
            for tag in ('b', 'bCs', 'i', 'iCs', 'u', 'rStyle'):
                for x in rp.findall(W(tag)):
                    rp.remove(x)
            r.append(rp)
        tt = OxmlElement('w:t'); tt.text = text
        tt.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(tt)
        return r
    h = mk(f'[{n}]')  # звичайний текст: Сергій прибрав із книжки всі гіперпосилання
    anchor = last if last.getparent().tag == W('p') else last.getparent()
    sp = mk(' ')
    anchor.addnext(sp)
    sp.addnext(h)
    if tail:
        h.addnext(mk(tail))


if __name__ == '__main__':
    main()
