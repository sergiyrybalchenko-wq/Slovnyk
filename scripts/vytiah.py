#!/usr/bin/env python3
"""Витягує статті з original/slovnyk-vykhidnyi.docx у dani/statti.tsv.

Кожна стаття — один або кілька абзаців тіла документа (індекси елементів body).
Стовпці: id, розділ (osn/abr), літера, елементи (індекси через кому), термін_uk, term_en, текст.
"""
import re, csv, sys, docx
from docx.table import Table
from docx.text.paragraph import Paragraph

SRC = 'original/slovnyk-vykhidnyi.docx'
DASH = r'\s*[-–—]\s*'
START = re.compile(r'^(?P<uk>[^A-Za-z]{1,140}?)' + DASH + r'(?P<en>[A-Za-z].*?)(?:' + DASH + r'|\s{2,}|$)')

def first_bold(p):
    for r in p.runs:
        if r.text.strip():
            return bool(r.bold)
    return False

def body_items(doc):
    """(індекс, тип, текст, абзац|None) для всіх елементів тіла."""
    out = []
    for i, el in enumerate(doc.element.body.iterchildren()):
        tag = el.tag.split('}')[1]
        if tag == 'p':
            p = Paragraph(el, doc)
            out.append((i, 'p', p.text, p))
        elif tag == 'tbl':
            t = Table(el, doc)
            out.append((i, 't', '\n'.join(c.text for row in t.rows for c in row.cells), None))
    return out

def extract(doc):
    items = body_items(doc)
    sec = None; letter = None; entries = []
    for i, kind, text, p in items:
        t = text.strip().replace(' ', ' ')
        if 'ОСНОВНА ЧАСТИНА' in t.replace('  ', ' ') and kind == 'p':
            sec = 'osn'; continue
        if t.startswith('РОЗШИФРОВКА СКЛАДНИХ'):
            sec = 'abr'; continue
        if t.startswith('АБЕТКОВИЙ ПОКАЗЧИК'):
            sec = None; continue
        if not sec or not t:
            continue
        if kind == 't':
            if len(t) > 40:
                entries.append(dict(sec=sec, letter=letter, els=[i], text=t, table=True))
            continue
        if re.fullmatch(r'[А-ЯІЇЄҐA-Z]', t):
            letter = t; continue
        is_start = (sec == 'abr') or first_bold(p) or START.match(t)
        if sec == 'osn' and not first_bold(p) and entries and not START.match(t):
            is_start = False
        if is_start or not entries:
            entries.append(dict(sec=sec, letter=letter, els=[i], text=t))
        else:
            entries[-1]['els'].append(i); entries[-1]['text'] += '\n' + t
    for n, e in enumerate(entries, 1):
        e['id'] = f'S{n:04d}'
        m = START.match(e['text'].split('\n')[0])
        e['uk'] = m.group('uk').strip() if m else e['text'].split(' - ')[0][:80].strip()
        e['en'] = m.group('en').strip() if m else ''
    return entries

if __name__ == '__main__':
    doc = docx.Document(SRC)
    E = extract(doc)
    with open('dani/statti.tsv', 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f, delimiter='\t')
        w.writerow(['id', 'розділ', 'літера', 'елементи', 'термін_uk', 'term_en', 'текст'])
        for e in E:
            w.writerow([e['id'], e['sec'], e['letter'] or '', ','.join(map(str, e['els'])), e['uk'], e['en'], e['text'].replace('\t', ' ')])
    print(len(E), 'статей;', sum(e['sec'] == 'abr' for e in E), 'у розділі абревіатур', file=sys.stderr)
