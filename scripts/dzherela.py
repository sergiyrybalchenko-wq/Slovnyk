#!/usr/bin/env python3
"""Будує dani/dzherela-novi.tsv — нові джерела для переліку «Використані джерела» (від № 36)
з колонок «джерело» і «url» файлу dani/novi.tsv.
Посилання на розділи чи сторінки однієї праці зводяться до одного запису (CANON).
Джерела вихідного словника (шифр Д…, а також NIST AI 100-2e2025 = Д35) не дублюються."""
import csv, re

MC = 'Міністерство цифрової трансформації України (2026). Словник термінів у сфері штучного інтелекту 2.0'
CANON = [  # (шаблон, текст запису, URL)
    (r'^Н08|Мінцифри \(2026\)\. Словник', MC, 'https://storage.thedigital.gov.ua/files/2/d2/9a599f3249bc1b1fb0597df31b9a3d2d.pdf'),
    (r'2024/1689|Artificial Intelligence Act', 'European Parliament and Council of the EU (2024). Regulation (EU) 2024/1689 laying down harmonised rules on artificial intelligence (Artificial Intelligence Act). Official Journal of the European Union', 'https://eur-lex.europa.eu/eli/reg/2024/1689/oj/eng'),
    (r'NIST AI 100-2e2025', None, ''),  # = Д35
    (r'Machine Learning Glossary', 'Google for Developers (2026). Machine Learning Glossary', 'https://developers.google.com/machine-learning/glossary'),
    (r'інтенсивний курс', 'Google for Developers (2026). Машинне навчання: інтенсивний курс (українська версія)', 'https://developers.google.com/machine-learning/crash-course?hl=uk'),
    (r'^Goodfellow I\., Bengio', 'Goodfellow, I., Bengio, Y. and Courville, A. (2016). Deep Learning. MIT Press', 'https://www.deeplearningbook.org/'),
    (r'^Jurafsky D\., Martin', 'Jurafsky, D. and Martin, J. H. (2026). Speech and Language Processing, 3rd ed. draft. Stanford University', 'https://web.stanford.edu/~jurafsky/slp3/'),
    (r'^Manning C\. D\., Raghavan', 'Manning, C. D., Raghavan, P. and Schütze, H. (2008). Introduction to Information Retrieval. Cambridge University Press', 'https://nlp.stanford.edu/IR-book/'),
    (r'Transformers documentation|Gradient accumulation|Gradient checkpointing|Continuous batching|Tensor parallelism for inference', 'Hugging Face (2026). Transformers documentation', ''),
    (r'Claude Platform Docs', 'Anthropic (2026). Claude Platform Docs', ''),
    (r'^MLflow\.', 'MLflow (2026). MLflow documentation', ''),
    (r'^scikit-learn developers', 'scikit-learn developers (2026). scikit-learn documentation', ''),
    (r'Attention Is All You Need', 'Vaswani, A. et al. (2017). Attention Is All You Need. NeurIPS 2017. arXiv:1706.03762', 'https://arxiv.org/abs/1706.03762'),
]

def canon(src, url):
    s = re.sub(r'^За матеріалами:\s*', '', src.strip()).split('; Huang Y.')[0].strip()
    if not s or s.startswith('Д'):
        return None, None
    for pat, text, u in CANON:
        if re.search(pat, s):
            return (text, (u or url.strip())) if text else (None, None)
    return s.rstrip('.'), url.strip()

def main():
    seen = {}
    for r in csv.DictReader(open('dani/novi.tsv', encoding='utf-8'), delimiter='\t'):
        t, u = canon(r['джерело'], r['url'])
        if t and t not in seen:
            seen[t] = u
    order = [MC] + sorted((k for k in seen if k != MC), key=str.lower)
    with open('dani/dzherela-novi.tsv', 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f, delimiter='\t', lineterminator='\n')
        w.writerow(['номер', 'текст'])
        for n, s in enumerate(order, 36):
            u = seen[s]
            w.writerow([n, f"{s}. [online] Available at: {u}." if u else s + '.'])
    print('нових джерел:', len(order))

if __name__ == '__main__':
    main()
