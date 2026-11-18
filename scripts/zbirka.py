#!/usr/bin/env python3
"""Збирає нову редакцію словника і звіт змін з вихідного DOCX і рішень.

Вхід:
  original/slovnyk-vykhidnyi.docx  — вихідний файл (не змінюється)
  dani/statti.tsv                  — статті вихідного файлу (scripts/vytiah.py)
  dani/rishennia.tsv               — рішення: id, дія (залишити/виправити/вилучити), новий_текст, причина, ...
  dani/novi.tsv                    — нові статті: термін_uk, term_en, визначення, причина, ...
  dani/dzherela-novi.tsv           — нові джерела для переліку в кінці (номер, текст)
Вихід:
  build/slovnyk-nova-redaktsiia.docx
  build/zminy.docx

Формат вихідного файлу зберігається: правка лише замінює, вилучає або вставляє абзаци,
нові абзаци копіюють форматування сусідньої статті. Титул, передмова, порядок розділів
і перелік джерел лишаються на місці; перелік джерел лише доповнюється.
"""
import copy, csv, os, re, sys
import docx
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(__file__))
from vytiah import extract, SRC  # noqa: E402

UA = "абвгґдеєжзиіїйклмнопрстуфхцчшщьюя"
DASH = ' – '


def sort_key(term):
    t = term.lower().replace('’', "'").replace('ʼ', "'")
    key = []
    for ch in t:
        if ch in UA:
            key.append((0, UA.index(ch)))
        elif ch.isalpha():
            key.append((1, ord(ch)))
        elif ch.isdigit():
            key.append((2, ord(ch)))
        elif ch == ' ':
            key.append((-1, 0))
    return key


def read_tsv(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding='utf-8') as f:
        return [r for r in csv.DictReader(f, delimiter='\t') if any((v or '').strip() for v in r.values())]


def set_runs(p, uk, en, definition, template_runs):
    """Переписує абзац статті: жирний термін – жирний англійський відповідник – визначення."""
    base = template_runs[0] if template_runs else None
    base_rpr = copy.deepcopy(base._element.rPr) if base is not None and base._element.rPr is not None else None
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    def add(text, bold):
        run = p.add_run(text)
        if base_rpr is not None:  # шрифт, розмір, мова — як у шаблонній статті
            run._element.insert(0, copy.deepcopy(base_rpr))
        run.bold = bold
        run.font.name = run.font.name or 'Times New Roman'
        return run
    add(uk, True)
    if en:
        add(DASH, False)
        add(en, True)
    add(DASH + definition, False)


def split_entry(text):
    """'Термін – English – визначення' -> (uk, en, def)."""
    parts = re.split(r'\s+[-–—]\s+', text.strip(), maxsplit=2)
    if len(parts) == 3:
        return parts
    if len(parts) == 2:
        return parts[0], '', parts[1]
    return text, '', ''


def main():
    os.makedirs('build', exist_ok=True)
    doc = docx.Document(SRC)
    body = list(doc.element.body.iterchildren())
    entries = {e['id']: e for e in extract(doc)}
    order = list(entries)
    dec = {r['id']: r for r in read_tsv('dani/rishennia.tsv')}
    new = read_tsv('dani/novi.tsv')
    changes = {'вилучено': [], 'відредаговано': [], 'додано': []}
    extra_terms = []

    # Шаблонний абзац для нових статей — перша стаття основної частини.
    first = entries[order[0]]
    tmpl_el = body[first['els'][0]]

    # Кожна стаття тримає посилання на свої елементи (а не індекси).
    for e in entries.values():
        e['refs'] = [body[i] for i in e['els']]
    moved = []

    def make_entry_par(uk, en, df):
        ne = copy.deepcopy(tmpl_el)
        p = Paragraph(ne, doc)
        set_runs(p, uk, en, df, p.runs)
        return ne

    # 1. Правки і вилучення.
    for sid in order:
        e = entries[sid]
        d = dec.get(sid)
        if not d:
            continue
        act = d['дія'].strip()
        if act == 'вилучити':
            for el in e['refs']:
                if el.getparent() is not None:
                    el.getparent().remove(el)
            changes['вилучено'].append((e['uk'], e['text'], d['причина']))
            e['deleted'] = True
        elif act in ('виправити', 'формат'):  # «формат» — лише впорядкування, у звіт не йде
            new_text = d['новий_текст'].strip()
            paras = [x.strip() for x in new_text.split('\\n') if x.strip()]
            uk, en, df = split_entry(paras[0])
            first_el = e['refs'][0]
            if first_el.tag.endswith('}p'):
                p0 = Paragraph(first_el, doc)
                set_runs(p0, uk, en, df, p0.runs)
            else:  # стаття в таблиці — замінюємо таблицю абзацом
                ne = make_entry_par(uk, en, df)
                first_el.addprevious(ne)
                first_el.getparent().remove(first_el)
                first_el = ne
            for el in e['refs'][1:]:
                if el.getparent() is not None:
                    el.getparent().remove(el)
            refs = [first_el]
            # Абзац, що починається з «§», — окрема стаття (розділення злитих статей).
            for para in paras[1:]:
                if para.startswith('§'):
                    u2, e2, d2 = split_entry(para[1:])
                    ne = make_entry_par(u2.strip(), e2.strip(), d2.strip())
                    extra_terms.append(u2.strip())
                else:
                    ne = copy.deepcopy(first_el)
                    pn = Paragraph(ne, doc)
                    for r in list(pn.runs):
                        r._element.getparent().remove(r._element)
                    pn.add_run(para)
                refs[-1].addnext(ne)
                refs.append(ne)
            e['refs'] = refs
            old_letter = (e['uk'].strip()[:1] or ' ').upper()
            e['uk'] = uk
            if e['sec'] == 'osn' and uk.strip()[:1].upper() != old_letter:
                moved.append(e)
            if act == 'виправити':
                changes['відредаговано'].append((uk, e['text'], new_text.replace('\\n', '\n').replace('§', ''), d['причина']))

    # 2. Нові статті і перейменовані на іншу літеру — за абеткою в межах своєї літери.
    alive = [entries[s] for s in order if entries[s]['sec'] == 'osn' and not entries[s].get('deleted') and entries[s] not in moved]

    def place(item):
        k = sort_key(item['uk'])
        L = item['uk'].strip()[:1].upper()
        same = [x for x in alive if x['uk'].strip()[:1].upper() == L]
        if same:
            after = None
            for x in same:
                if sort_key(x['uk']) <= k:
                    after = x
            if after is None:
                for el in item['refs']:
                    same[0]['refs'][0].addprevious(el)
                alive.insert(alive.index(same[0]), item)
            else:
                anchor = after['refs'][-1]
                for el in item['refs']:
                    anchor.addnext(el); anchor = el
                alive.insert(alive.index(after) + 1, item)
        else:  # немає статей цієї літери — після останньої статті меншої літери
            prev = [x for x in alive if sort_key(x['uk']) <= k]
            anchor = (prev[-1] if prev else alive[0])['refs'][-1]
            for el in item['refs']:
                anchor.addnext(el); anchor = el
            alive.insert(alive.index(prev[-1]) + 1 if prev else 0, item)

    for e in moved:
        for el in e['refs']:
            el.getparent().remove(el)
        place(e)
    for n in sorted(new, key=lambda r: sort_key(r['термін_uk'])):
        uk = n['термін_uk'].strip()
        item = dict(uk=uk, sec='osn', refs=[make_entry_par(uk, n['term_en'].strip(), n['визначення'].strip())])
        place(item)
        changes['додано'].append((uk, f"{uk} – {n['term_en'].strip()} – {n['визначення'].strip()}", n['причина']))

    # 3. Чернеткові порожнечі: між статтями основної частини і розділу абревіатур — без порожніх абзаців.
    paras = doc.paragraphs
    start = next(i for i, p in enumerate(paras) if 'ОСНОВНА' in p.text.replace(' ', ' ') and 'ЧАСТИНА' in p.text and i > 100)
    idx_start = next(i for i, p in enumerate(paras) if p.text.strip().startswith('АБЕТКОВИЙ ПОКАЗЧИК'))
    for p in paras[start + 1:idx_start]:
        if not p.text.strip() and not p._element.xpath('.//w:drawing|.//w:pict|.//w:br[@w:type="page"]'):
            nxt = p._element.getnext()
            # одна порожня лінія перед заголовком літери лишається
            if nxt is not None and nxt.tag.endswith('}p') and re.fullmatch(r'[А-ЯІЇЄҐA-Z]', Paragraph(nxt, doc).text.strip()):
                continue
            p._element.getparent().remove(p._element)

    # 4. Абетковий покажчик: повний перелік термінів основної частини.
    paras = doc.paragraphs
    idx_start = next(i for i, p in enumerate(paras) if p.text.strip().startswith('АБЕТКОВИЙ ПОКАЗЧИК'))
    src_start = next(i for i, p in enumerate(paras) if p.text.strip().startswith('ВИКОРИСТАНІ ДЖЕРЕЛА'))
    old = paras[idx_start + 1:src_start]
    letter_tmpl = next(p for p in old if re.fullmatch(r'[А-Я]', p.text.strip()))._element
    term_tmpl = next(p for p in old if len(p.text.strip()) > 2)._element
    blank_tmpl = next((p for p in old if not p.text.strip()), None)
    anchor = paras[idx_start]._element
    for p in old:
        p._element.getparent().remove(p._element)
    terms = sorted({re.sub(r'\s+', ' ', e['uk']).strip() for e in alive} | set(extra_terms), key=sort_key)
    cur = None
    for t in terms:
        L = t[:1].upper()
        if L != cur:
            if blank_tmpl is not None:
                b = copy.deepcopy(blank_tmpl._element); anchor.addnext(b); anchor = b
            le = copy.deepcopy(letter_tmpl); anchor.addnext(le); anchor = le
            Paragraph(le, doc).runs[0].text = L
            for r in Paragraph(le, doc).runs[1:]:
                r._element.getparent().remove(r._element)
            cur = L
        te = copy.deepcopy(term_tmpl); anchor.addnext(te); anchor = te
        pt = Paragraph(te, doc)
        pt.runs[0].text = t
        for r in pt.runs[1:]:
            r._element.getparent().remove(r._element)

    # 5. Перелік джерел: нові джерела після останнього нумерованого.
    paras = doc.paragraphs
    numbered = [p for p in paras if re.match(r'\s*\d+\.', p.text) and p.paragraph_format.first_line_indent == paras[-1].paragraph_format.first_line_indent]
    last_src = numbered[-1]._element if numbered else paras[-1]._element
    for s in read_tsv('dani/dzherela-novi.tsv'):
        ne = copy.deepcopy(last_src); last_src.addnext(ne); last_src = ne
        p = Paragraph(ne, doc)
        p.runs[0].text = f"{s['номер']}.{s['текст']}"
        for r in p.runs[1:]:
            r._element.getparent().remove(r._element)

    # 6. Єдиний шрифт статей: чернеткові вставки (Helvetica, Calibri тощо) -> Times New Roman.
    paras = doc.paragraphs
    for p in paras[start + 1:]:
        for r in p.runs:
            if r.font.name is None or ';' in r.font.name or r.font.name != 'Times New Roman':
                r.font.name = 'Times New Roman'
                rf = r._element.rPr.rFonts
                rf.set(docx.oxml.ns.qn('w:cs'), 'Times New Roman')
                rf.set(docx.oxml.ns.qn('w:eastAsia'), 'Times New Roman')

    doc.save('build/slovnyk-nova-redaktsiia.docx')
    zvit(changes)
    print({k: len(v) for k, v in changes.items()}, 'термінів у покажчику:', len(terms))


def zvit(changes):
    z = docx.Document()
    st = z.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = docx.shared.Pt(12)
    z.add_heading('Зміни в новій редакції словника «Штучний інтелект»', 0)
    z.add_paragraph('Документ перелічує статті, які вилучено, відредаговано і додано, з коротким поясненням кожного рішення.')
    titles = {'вилучено': '1. Вилучені статті', 'відредаговано': '2. Відредаговані статті', 'додано': '3. Додані статті'}
    for key in ('вилучено', 'відредаговано', 'додано'):
        items = sorted(changes[key], key=lambda x: sort_key(x[0]))
        z.add_heading(f'{titles[key]} ({len(items)})', 1)
        for n, it in enumerate(items, 1):
            p = z.add_paragraph()
            p.add_run(f'{n}. {it[0]}').bold = True
            if key == 'відредаговано':
                z.add_paragraph('Було: ' + it[1].replace('\n', ' '))
                z.add_paragraph('Стало: ' + it[2].replace('\n', ' '))
            else:
                z.add_paragraph(it[1].replace('\n', ' '))
            r = z.add_paragraph().add_run('Чому: ' + it[-1]); r.italic = True
    z.save('build/zminy.docx')


if __name__ == '__main__':
    main()
