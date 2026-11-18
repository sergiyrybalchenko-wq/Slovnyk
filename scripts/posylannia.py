#!/usr/bin/env python3
"""Генерує dani/posylannia.tsv: біля яких визначень можна поставити посилання на джерело.
Джерела: збіги з Мінцифри (dani/mintsyfra-zistavlennia.tsv), нові статті (dani/novi.tsv),
авторські визначення з установленим джерелом (dani/rishennia.tsv, колонка «джерело»)."""
import csv
MC = 'https://storage.thedigital.gov.ua/files/2/d2/9a599f3249bc1b1fb0597df31b9a3d2d.pdf'
# URL відомих авторських джерел: джерело -> (url, стан перевірки)
URLS = {
    'Moveworks AI Terms Glossary — Adapters': ('https://moveworks.com/us/en/resources/ai-terms-glossary/adapters', 'перевірено рецензентом 18.11.2026'),
    'USAID Digital Strategy 2020–2024 (Digital Ecosystem Framework)': ('https://uatweb.usaid.gov/sites/default/files/2022-05/Digital_Strategy_Digital_Ecosystem_Final.pdf', 'перевірено рецензентом 18.11.2026'),
    'DashboardFox BI Glossary — NLG (корпоративний глосарій)': ('https://dashboardfox.com/business-intelligence-glossary/natural-language-generation-nlg/', 'перевірено рецензентом 18.11.2026'),
    'NIST AI 100-2': ('https://csrc.nist.gov/pubs/ai/100/2/e2025/final', 'перевірено рецензентом 18.11.2026'),
}
rd = lambda p: list(csv.DictReader(open(p, encoding='utf-8'), delimiter='\t'))
dec = {r['id']: r for r in rd('dani/rishennia.tsv')}
S = {r['id']: r for r in rd('dani/statti.tsv')}
out = []
for r in rd('dani/mintsyfra-zistavlennia.tsv'):
    i = r['id']
    if i and r['збіг'] in ('дослівно', 'близько') and dec.get(i, {}).get('дія') != 'вилучити':
        out.append([f"{i} {S[i]['термін_uk'].strip()}", f"Н08 Мінцифри, Словник термінів у сфері ШІ 2.0, № {r['номер_мц']} ({r['збіг']})", MC, 'після визначення', 'URL відкривається 18.11.2026'])
for r in rd('dani/novi.tsv'):
    out.append([f"нова: {r['термін_uk']}", r['джерело'], r['url'], 'після визначення', r['перевірка']])
for i, d in sorted(dec.items()):
    s = d['джерело']
    if d['дія'] != 'вилучити' and s and s != 'Н08 (термін)' and 'Мінцифр' not in s:
        url, st = URLS.get(s, ('', 'URL ще не перевірено'))
        out.append([f"{i} {S[i]['термін_uk'].strip()}", s, url, 'після визначення', st])
with open('dani/posylannia.tsv', 'w', encoding='utf-8', newline='') as f:
    w = csv.writer(f, delimiter='\t', lineterminator='\n')
    w.writerow(['id_або_термін', 'джерело', 'url', 'місце', 'перевірено'])
    w.writerows(out)
print('посилань:', len(out))
