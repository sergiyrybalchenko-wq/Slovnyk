#!/usr/bin/env python3
"""Будує dani/storinky.tsv — точне місце в джерелі для посилань у тексті: «[36, с. 12]».
Ставимо лише там, де місце встановлено надійно:
  • Мінцифри, Словник термінів у сфері ШІ 2.0 (PDF, 37 с.): таблиця по 5 термінів на сторінку
    з с. 4 (№ 1–5) до с. 30 (№ 131–135), далі с. 31 — № 136–141, с. 32 — № 142–146.
    Звірено з PDF за координатами номерів і вибірково очима (с. 22, 23, 29–32), 20.11.2026.
  • NIST (друковані номери сторінок; знайдено визначення в тексті PDF 20.11.2026) — таблиця NIST нижче.
  • Регламент (ЄС) 2024/1689 і 2022/2065 — стаття й пункт (для нормативних актів замість сторінки).
Інші джерела (вебсторінки, глосарії, книжки у вебверсії, препринти) — без сторінок."""
import csv, re

def mc_page(n):
    if n <= 135:
        return 4 + (n - 1) // 5
    return 31 if n <= 141 else 32

NIST = {  # термін (нова стаття) -> сторінка за друкованою нумерацією
    'Толерантність до ризику': 'p. 7', "Учасники у сфері ШІ (суб'єкти ШІ)": 'p. 2', 'Залишковий ризик': 'p. 8',
    'Отруєння моделі': 'p. 26', 'Бекдор-атака (атака із закладкою)': 'p. 22', 'Отруєння з чистими мітками': 'p. 21',
    'Цільове отруєння': 'p. 21', 'Атака виведення властивостей': 'p. 31',
    'Синтетичний контент': 'p. 74', 'Цифровий водяний знак': 'p. 5', 'Автентифікація контенту': 'p. 73',
    'Дезінформація': 'p. 10', 'Хибна інформація (місінформація)': 'p. 10',
}

def art(src):
    m = re.search(r'(?:ст\.|Article) (\d+)((?:\([^)]+\))*)(, (?:ст\. \d+|додаток [IVX]+))?', src)
    if not m:
        return None
    parts = re.findall(r'\(([^)]+)\)', m.group(2))
    s = f'ст. {m.group(1)}'
    if len(parts) == 1:
        s += f', п. {parts[0]}'
    elif len(parts) == 2:
        s += f', ч. {parts[0]}, п. {parts[1]}'
    if m.group(3):
        s += m.group(3).replace('додаток', 'дод.').replace(', ст.', '; ст.')
    return s

T = {r['термін_uk'].strip().lower(): int(r['номер']) for r in csv.DictReader(open('dani/mintsyfra.tsv', encoding='utf-8'), delimiter='\t')}
novi = {r['термін_uk'].strip(): r for r in csv.DictReader(open('dani/novi.tsv', encoding='utf-8'), delimiter='\t')}
out = []
for r in csv.DictReader(open('dani/posylannia.tsv', encoding='utf-8'), delimiter='\t'):
    k, src = r['id_або_термін'], r['джерело']
    loc = None
    m = re.search(r'Мінцифри, Словник термінів у сфері ШІ 2\.0, № (\d+)', src)
    if m:
        loc = f'с. {mc_page(int(m.group(1)))}'
    elif k.startswith('нова: '):
        t = k[6:].strip()
        if src.startswith('Н08') or 'Мінцифри (2026). Словник' in src:
            n = T.get(t.lower())
            loc = f'с. {mc_page(n)}' if n else None
        elif t in NIST:
            loc = NIST[t]
        elif '2024/1689' in src or '2022/2065' in src:
            loc = art(src)
    if loc:
        out.append([k, loc])
with open('dani/storinky.tsv', 'w', encoding='utf-8', newline='') as f:
    w = csv.writer(f, delimiter='\t', lineterminator='\n')
    w.writerow(['id_або_термін', 'місце'])
    w.writerows(out)
print('місць:', len(out))
