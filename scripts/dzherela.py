#!/usr/bin/env python3
"""Будує dani/dzherela-novi.tsv — нові джерела для переліку «Використані джерела» (від № 36)
з колонок «джерело» і «url» файлу dani/novi.tsv.
Посилання на розділи чи сторінки однієї праці зводяться до одного запису (CANON).
Джерела вихідного словника (шифр Д…, а також NIST AI 100-2e2025 = Д35) не дублюються."""
import csv, re

MC = 'Словник термінів у сфері штучного інтелекту 2.0 / Міністерство цифрової трансформації України. Київ, 2026. 37 с.'
CANON = [  # (шаблон, текст запису, URL)
    (r'^Н08|Мінцифри \(2026\)\. Словник', MC, 'https://storage.thedigital.gov.ua/files/2/d2/9a599f3249bc1b1fb0597df31b9a3d2d.pdf'),
    (r'2024/1689|Artificial Intelligence Act', 'Regulation (EU) 2024/1689 of the European Parliament and of the Council of 13 June 2024 laying down harmonised rules on artificial intelligence (Artificial Intelligence Act). Official Journal of the European Union. 2024. L 2024/1689', 'https://eur-lex.europa.eu/eli/reg/2024/1689/oj/eng'),
    (r'CSRC Glossary|Computer Security Resource Center Glossary|Computer Security Resource Center\. National',
     'Glossary. Computer Security Resource Center. National Institute of Standards and Technology', 'https://csrc.nist.gov/glossary'),
    (r'NIST AI 100-2e2025', None, ''),  # = Д35
    (r'NIST AI 600-1', None, ''),  # = Д24
    (r'Framework Convention on Artificial Intelligence', None, ''),  # = Д06
    (r'General-Purpose AI Code of Practice', None, ''),  # = Д09
    (r'Machine Learning Glossary', 'Machine Learning Glossary. Google for Developers. 2026', 'https://developers.google.com/machine-learning/glossary'),
    (r'інтенсивний курс', 'Машинне навчання: інтенсивний курс. Google for Developers. 2026', 'https://developers.google.com/machine-learning/crash-course?hl=uk'),
    (r'^Goodfellow I\., Bengio', 'Goodfellow I., Bengio Y., Courville A. Deep Learning. Cambridge, MA : MIT Press, 2016', 'https://www.deeplearningbook.org/'),
    (r'^Jurafsky D\., Martin', 'Jurafsky D., Martin J. H. Speech and Language Processing. 3rd ed. draft. Stanford University, 2026', 'https://web.stanford.edu/~jurafsky/slp3/'),
    (r'^Manning C\. D\., Raghavan', 'Manning C. D., Raghavan P., Schütze H. Introduction to Information Retrieval. Cambridge : Cambridge University Press, 2008', 'https://nlp.stanford.edu/IR-book/'),
    (r'Transformers documentation|Gradient accumulation|Gradient checkpointing|Continuous batching|Tensor parallelism for inference', 'Transformers documentation. Hugging Face. 2026', 'https://huggingface.co/docs/transformers/index'),
    (r'Claude Platform Docs', 'Claude Platform Docs. Anthropic. 2026', 'https://platform.claude.com/docs/en/home'),
    (r'^MLflow\.', 'MLflow documentation. 2026', 'https://mlflow.org/docs/latest/'),
    (r'^scikit-learn developers', 'scikit-learn documentation. scikit-learn developers. 2026', 'https://scikit-learn.org/stable/'),
    (r'Attention Is All You Need', 'Vaswani A., Shazeer N., Parmar N. et al. Attention Is All You Need. Advances in Neural Information Processing Systems 30 (NeurIPS 2017). arXiv:1706.03762', 'https://arxiv.org/abs/1706.03762'),
]

# Джерела авторських визначень, які ставимо посиланням у тексті (dani/posylannia.tsv).
EXTRA = [
    ('USAID. Digital Strategy 2020–2024: Digital Ecosystem Framework. Washington, 2020', 'https://uatweb.usaid.gov/sites/default/files/2022-05/Digital_Strategy_Digital_Ecosystem_Final.pdf'),
]


# Джерела вихідного переліку (№ 1–35), на які посилаються нові статті.
D_NUM = [(r'^Д(\d+)', None), (r'NIST AI 100-2e2025', 35), (r'NIST AI 600-1', 24),
         (r'Framework Convention on Artificial Intelligence', 6), (r'General-Purpose AI Code of Practice', 9)]


def d_number(src):
    for pat, n in D_NUM:
        m = re.search(pat, src.strip())
        if m:
            return n or int(m.group(1))
    return None


def canon(src, url):
    s = re.sub(r'^За матеріалами:\s*', '', src.strip()).split('; Huang Y.')[0].strip()
    if not s or s.startswith('Д'):
        return None, None
    for pat, text, u in CANON:
        if re.search(pat, s):
            return (text, (u or url.strip())) if text else (None, None)
    s = re.sub(r'(?<!\.)\.\.(?!\.)', '.', s)  # «et al..» -> «et al.»
    s = re.sub(r', п\. [\d.]+(?= \()', '', s)
    s = re.sub(r', ст\. \d+\([a-z0-9]+\)', '', s)  # місце в акті — у посиланні, не в описі джерела  # пункт стосується однієї статті, не праці
    return s.rstrip('.'), url.strip()

def dstu(s, u):
    """Опис за ДСТУ 8302:2015: «… URL: … (дата звернення: …).» або «… DOI: ….»"""
    s = s.rstrip('.').replace(' [online]', '')
    s = re.sub(r"\b([A-Z][\w’'\-]+), ((?:[A-Z]\.(?:[ -]?[A-Z]\.)*))(?=[,. ])", r'\1 \2', s)  # «Hinton, G.» -> «Hinton G.»
    s = re.sub(r'\s*\((?:картка|голова)[^)]*\)', '', s)
    s = re.sub(r',?\s*дата публікації не вказана(?: \([^)]*\))?', '', s)
    s = re.sub(r'\s*\((?:Hugging Face|Google Research|Google Cloud|IBM)\)', '', s)  # платформа вже є після назви
    s = re.sub(r'(?<!\.)\.\.(?!\.)', '.', s)
    s = re.sub(r'\b([A-Z])-([A-Z])\.', r'\1.-\2.', s)  # «M-T.» -> «M.-T.»
    m = re.search(r'arxiv\.org/(?:abs|pdf)/(\d{4}\.\d{4,5})', u or '')
    if m and 'arXiv:' not in s:
        s = s.replace('arXiv,', f'arXiv:{m.group(1)},', 1) if 'arXiv,' in s else f'{s}. arXiv:{m.group(1)}'
    # Області опису за ДСТУ 8302:2015 відокремлюємо крапкою: «…, arXiv:…, 2015» -> «…. arXiv:…. 2015».
    s = re.sub(r',\s*(arXiv:[\d.]+(?:v\d)?)', r'. \1', s)
    s = re.sub(r',\s*((?:19|20)\d{2})$', r'. \1', s)
    if not u:
        return s + '.'
    if 'doi.org/' in u:
        return f'{s}. DOI: {u}.'
    return f'{s}. URL: {u} (дата звернення: 20.11.2026).'


def main():
    seen = {}
    for r in csv.DictReader(open('dani/novi.tsv', encoding='utf-8'), delimiter='\t'):
        t, u = canon(r['джерело'], r['url'])
        if t and t not in seen:
            seen[t] = u
    for t, u in EXTRA:
        seen.setdefault(t, u)
    try:  # джерела авторських визначень (scripts/avtor_dzherela.py)
        for r in csv.DictReader(open('dani/posylannia-avtor.tsv', encoding='utf-8'), delimiter='\t'):
            t, u = canon(r['джерело'], r['url'])
            if t and t not in seen:
                seen[t] = u
    except FileNotFoundError:
        pass
    order = [MC] + sorted((k for k in seen if k != MC), key=str.lower)
    with open('dani/dzherela-novi.tsv', 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f, delimiter='\t', lineterminator='\n')
        w.writerow(['номер', 'ключ', 'текст'])
        for n, s in enumerate(order, 36):
            w.writerow([n, s, dstu(s, seen[s])])
    print('нових джерел:', len(order))


if __name__ == '__main__':
    main()
