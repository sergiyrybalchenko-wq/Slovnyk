#!/usr/bin/env python3
"""Зводить знайдені джерела авторських визначень (партії пошуку r*.tsv + вердикт рецензента)
у dani/posylannia-avtor.tsv: термін, збіг, джерело (канонічний опис за ДСТУ 8302:2015), url, місце, доказ, перевірка.
Запуск: python3 scripts/avtor_dzherela.py dani/poshuk"""
import csv, glob, re, sys, urllib.parse

D = '(дата звернення: 20.11.2026)'
WIKI = 'Wikipedia : the free encyclopedia'
CANON = [
    (r'Glossary_of_artificial_intelligence', 'Glossary of artificial intelligence. ' + WIKI,
     'https://en.wikipedia.org/wiki/Glossary_of_artificial_intelligence'),
    (r'Glossary_of_computer_science', 'Glossary of computer science. ' + WIKI,
     'https://en.wikipedia.org/wiki/Glossary_of_computer_science'),
    (r'NIST\.IR\.8269', 'Tabassi E., Burns K. J., Hadjimichael M., Molina-Markham A. D., Sexton J. T. A Taxonomy and '
     'Terminology of Adversarial Machine Learning : Draft NISTIR 8269. Gaithersburg : National Institute of Standards '
     'and Technology, 2019', 'https://doi.org/10.6028/NIST.IR.8269-draft'),
    (r'NIST\.IR\.8312', 'Phillips P. J., Hahn C. A., Fontana P. C., Broniatowski D. A., Przybocki M. A. Four Principles '
     'of Explainable Artificial Intelligence : Draft NISTIR 8312. Gaithersburg : National Institute of Standards and '
     'Technology, 2020', 'https://nvlpubs.nist.gov/nistpubs/ir/2020/NIST.IR.8312-draft.pdf'),
    (r'csrc\.nist\.gov/glossary', 'Glossary. Computer Security Resource Center. National Institute of Standards and '
     'Technology', 'https://csrc.nist.gov/glossary'),
    (r'developers\.google\.com/machine-learning/glossary', 'Machine Learning Glossary. Google for Developers. 2026',
     'https://developers.google.com/machine-learning/glossary'),
]


# Статті-відсилання («Див. …»): посилання на джерело нічого не дає.
SKIP = {"Комп'ютерне прослуховування"}


def canon(url):
    for pat, text, u in CANON:
        if re.search(pat, url):
            return text, u
    m = re.search(r'wikipedia\.org/(?:wiki/|w/index\.php\?title=)([^&#]+)', url)
    if m:
        title = urllib.parse.unquote(m.group(1)).replace('_', ' ')
        wiki = 'Вікіпедія : вільна енциклопедія' if 'uk.wikipedia' in url else WIKI
        return f'{title}. {wiki}', url
    raise ValueError(url)


def loc(s):
    m = re.search(r'p+\.\s*([\d–-]+)', s or '')
    return f'p. {m.group(1).replace("-", "–")}' if m else ''


def main(d):
    ver = {}
    vf = d + '/verdykt.tsv'
    for r in csv.DictReader(open(vf, encoding='utf-8'), delimiter='\t'):
        ver[r['термін'].strip()] = (r['вердикт'].strip(), r.get('коментар', ''))
    out = []
    for f in sorted(glob.glob(d + '/r[0-9].tsv')):
        for r in csv.DictReader(open(f, encoding='utf-8'), delimiter='\t'):
            t = r['термін'].strip()
            v, com = ver.get(t, ('', ''))
            if v.startswith('не згоден') or not v:
                continue
            if t in SKIP:
                continue
            text, url = canon(r['url'])
            zb = 'частково' if v.startswith('частково') else r['збіг'].strip()
            out.append([t, zb, text, url, '' if '8312' in url else loc(r['місце']), r['доказ'], f"{r['перевірено']}; рецензент: {v}" + (f' ({com})' if com else '')])
    with open('dani/posylannia-avtor.tsv', 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f, delimiter='\t', lineterminator='\n')
        w.writerow(['термін', 'збіг', 'джерело', 'url', 'місце', 'доказ', 'перевірка'])
        w.writerows(out)
    print('рядків:', len(out))


if __name__ == '__main__':
    main(sys.argv[1])
