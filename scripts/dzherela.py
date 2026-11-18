#!/usr/bin/env python3
"""Будує dani/dzherela-novi.tsv (нові джерела для переліку «Використані джерела», від № 36)
з колонок «джерело» і «url» файлу dani/novi.tsv. Одне джерело — один номер;
якщо джерело має кілька URL (різні сторінки), береться перший."""
import csv
rows = list(csv.DictReader(open('dani/novi.tsv', encoding='utf-8'), delimiter='\t'))
seen = {}
for r in rows:
    s = r['джерело'].strip()
    if s.startswith('Н08'):
        s = 'Міністерство цифрової трансформації України (2026). Словник термінів у сфері штучного інтелекту 2.0'
    if s and s not in seen and not s.startswith('Д'):
        seen[s] = r['url'].strip()
with open('dani/dzherela-novi.tsv', 'w', encoding='utf-8', newline='') as f:
    w = csv.writer(f, delimiter='\t', lineterminator='\n')
    w.writerow(['номер', 'текст'])
    mc = [k for k in seen if k.startswith('Міністерство цифрової')]
    order = mc + sorted((k for k in seen if k not in mc), key=str.lower)
    for n, s in enumerate(order, 36):
        u = seen[s]
        w.writerow([n, f"{s.rstrip('.')}. [online] Available at: {u}." if u else s])
print('нових джерел:', len(seen))
