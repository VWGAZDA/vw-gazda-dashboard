#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Wyciaga Informacje Produkcyjna (zalacznik do Informacji Handlowej VW) z PDF
i zapisuje ja w formacie sekcji `produkcja` obiektu DATA dashboardu.

Uzycie:
  python3 parse_ih.py --dir <folder z pobranymi> [--out produkcja.json]
  python3 parse_ih.py --pdf <plik.pdf>        [--out produkcja.json]

Wynik (JSON):
{
  "ih_numer": "IH 40/2026",
  "ih_tydzien": "31/2026",
  "ih_data": "27 lipca 2026",
  "ih_plik": "Informacja handlowa 40_2026 ... .pdf",
  "produkcja": {"komunikaty": [...], "przerwy": [...], "kalendarz": [...], "restrykcje": [...]}
}
Wpisy: {"tytul": str, "data": str, "tresc": str} lub {"tytul": str, "tabela": [[...]]}
       opcjonalnie "pre": true - zachowaj lamania linii.
"""
import argparse, glob, json, os, re, sys

import pdfplumber

BULLET = '▪'          # ▪
NBSP = ' '


def clean(s):
    return re.sub(r'[ \t]+', ' ', (s or '').replace(NBSP, ' ')).strip()


def find_pdf(directory):
    pats = ['*nformacja*produkcyjna*.pdf', '*IH*produkcyjna*.pdf', '*nformacja*handlowa*.pdf']
    found = []
    for pat in pats:
        for f in glob.glob(os.path.join(directory, pat)):
            found.append(f)
        for f in glob.glob(os.path.join(directory, '**', pat), recursive=True):
            found.append(f)
    found = sorted(set(found), key=lambda f: os.path.getmtime(f), reverse=True)
    return found[0] if found else None


def page_lines(pg, tol=3):
    """Zwraca liste wierszy: [(top, [words...])] posortowana po y."""
    d = {}
    for w in pg.extract_words():
        d.setdefault(round(w['top'] / tol), []).append(w)
    return [(k * tol, sorted(v, key=lambda w: w['x0'])) for k, v in sorted(d.items())]


def is_footer(txt):
    t = txt.lower()
    return (t.startswith('informacja handlowa') and 'csd' in t) or t.startswith('confidential') \
        or t.startswith('słownik pojęć') or t.startswith('*podane') or t.startswith('*występowanie') \
        or t.startswith('bazarestrykcji') or t.startswith('handlowych.') or t.startswith('należy uwzględnić')


def parse(pdf_path):
    out = {'ih_numer': '', 'ih_tydzien': '', 'ih_data': '', 'ih_plik': os.path.basename(pdf_path),
           'produkcja': {'komunikaty': [], 'przerwy': [], 'kalendarz': [], 'restrykcje': []}}
    with pdfplumber.open(pdf_path) as pdf:
        texts = [(pg.extract_text() or '') for pg in pdf.pages]
        joined = '\n'.join(texts)

        m = re.search(r'Informacja\s+Handlowa\s+(\d+)\s*/\s*(\d{4})', joined)
        if m:
            out['ih_numer'] = 'IH %s/%s' % (m.group(1), m.group(2))
        m = re.search(r'na\s+tydzie[nń]\s+(\d+)\s*/\s*(\d{4})', joined)
        if m:
            out['ih_tydzien'] = '%s/%s' % (m.group(1), m.group(2))
        m = re.search(r'\|\s*(\d{1,2}\s+\w+\s+\d{4})\s*\|', joined)
        if m:
            out['ih_data'] = m.group(1)

        naglowek = '%s · Informacja Produkcyjna na tydzień %s' % (
            out['ih_numer'] or 'IH', out['ih_tydzien'] or '—')

        # ---------------- KOMUNIKATY (punktory) ----------------------
        kom = []
        for pi, txt in enumerate(texts):
            head = clean(txt.split('\n')[1] if len(txt.split('\n')) > 1 else '')
            if head.upper() != 'KOMUNIKATY':
                continue
            body = [l for l in txt.split('\n')[2:] if not is_footer(clean(l))]
            cur = None
            for line in body:
                l = clean(line)
                if not l:
                    continue
                if l.startswith(BULLET):
                    if cur:
                        kom.append(cur)
                    l = clean(l.lstrip(BULLET))
                    # tytul = fragment przed pierwszym dwukropkiem, max 90 znakow
                    tytul = l.split(':')[0] if ':' in l[:90] else l[:90]
                    cur = {'tytul': clean(tytul), 'data': naglowek,
                           'tresc': clean(l[len(tytul):].lstrip(': ')) or clean(l)}
                elif cur:
                    cur['tresc'] = clean(cur['tresc'] + ' ' + l)
                else:
                    cur = {'tytul': l[:90], 'data': naglowek, 'tresc': l}
            if cur:
                kom.append(cur)
        out['produkcja']['komunikaty'] = kom

        # ---------------- KALENDARZ PRODUKCYJNY ----------------------
        for pi, txt in enumerate(texts):
            # naglowek wlasnej strony (nie spis tresci)
            if not re.match(r'^\s*\d+\s*\n\s*KALENDARZ PRODUKCYJNY', txt, re.I):
                continue
            pg = pdf.pages[pi]
            # granice kolumn wyznaczone z ukladu tabeli VW (szerokosc strony 960)
            sc = pg.width / 960.0
            bounds = [200 * sc, 350 * sc, 620 * sc, 760 * sc]

            def col(w):
                for j, b in enumerate(bounds):
                    if w['x0'] < b:
                        return j
                return 4

            rows, pending = [], None
            for top, ws in page_lines(pg):
                cells = ['', '', '', '', '']
                for w in ws:
                    j = col(w)
                    cells[j] = (cells[j] + ' ' + w['text']).strip()
                line = clean(' '.join(cells))
                if not line or is_footer(line):
                    continue
                up = line.upper()
                if up.startswith(('KALENDARZ', 'MODEL', 'NAJBLIŻSZY', 'TYDZIEŃ')) or up == str(pi + 1):
                    continue
                has_model = bool(cells[0]) and re.search(r'\([A-Z0-9]{2,4}\)', cells[0])
                has_vals = any(cells[1:])
                if has_model and has_vals:
                    rows.append(cells)
                elif has_vals and not cells[0]:
                    pending = cells
                elif has_model and pending:
                    pending[0] = cells[0]
                    rows.append(pending)
                    pending = None
            if rows:
                out['produkcja']['kalendarz'] = [{
                    'tytul': 'Kalendarz produkcyjny — tydzień %s' % (out['ih_tydzien'] or '—'),
                    'data': naglowek,
                    'tabela': [['Model', 'Najbliższy zatw. tydzień', 'Najwcześniejszy m-c dla nowych komisji',
                                'Opóźnienie prod.', 'Dostawa fabryka → VGP']] + rows,
                    'tresc': 'Terminy orientacyjne — mogą się wydłużyć. Uwzględnij też restrykcje.',
                }]
            break

        # ---------------- RESTRYKCJE (wiernie, bez interpretacji) ----
        for pi, txt in enumerate(texts):
            if not re.match(r'^\s*\d+\s*\n\s*RESTRYKCJE', txt, re.I):
                continue
            body = [clean(l) for l in txt.split('\n')[2:]]
            body = [l for l in body if l and not is_footer(l)
                    and l.upper() != 'MODEL KOD OPIS RESTRYKCJI']
            if body:
                out['produkcja']['restrykcje'] = [{
                    'tytul': 'Elementy restrykcyjne — %s' % (out['ih_tydzien'] or '—'),
                    'data': naglowek, 'pre': True,
                    'tresc': '\n'.join(body) +
                             '\n\nWystąpienie któregokolwiek z powyższych elementów w specyfikacji '
                             'może wydłużyć czas zatwierdzenia zamówienia do produkcji, a tym samym '
                             'czas dostawy do Klienta.',
                }]
            break

        # ---------------- PRZERWY PRODUKCYJNE ------------------------
        prz = []
        for pi, txt in enumerate(texts):
            if 'PRZERW' not in txt.upper():
                continue
            body = [clean(l) for l in txt.split('\n')[1:]]
            body = [l for l in body if l and not is_footer(l)]
            if body:
                prz.append({'tytul': 'Przerwy produkcyjne', 'data': naglowek,
                            'pre': True, 'tresc': '\n'.join(body)})
            break
        if not prz:
            prz = [{'tytul': 'Brak informacji o przerwach', 'data': naglowek,
                    'tresc': 'Załącznik %s nie zawiera sekcji o przerwach produkcyjnych.'
                             % (out['ih_numer'] or 'IH')}]
        out['produkcja']['przerwy'] = prz

    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dir')
    ap.add_argument('--pdf')
    ap.add_argument('--out', default='produkcja.json')
    a = ap.parse_args()

    path = a.pdf
    if not path:
        if not a.dir:
            sys.exit('BLAD: podaj --pdf albo --dir')
        path = find_pdf(a.dir)
        if not path:
            sys.exit('BLAD: nie znaleziono PDF-a z Informacja Produkcyjna w %s' % a.dir)

    res = parse(path)
    open(a.out, 'w', encoding='utf-8').write(json.dumps(res, ensure_ascii=False, indent=1))
    p = res['produkcja']
    print('OK  %s' % a.out)
    print('    plik           %s' % res['ih_plik'])
    print('    %s · tydzień %s · %s' % (res['ih_numer'], res['ih_tydzien'], res['ih_data']))
    for k in ('komunikaty', 'przerwy', 'kalendarz', 'restrykcje'):
        print('    %-14s %d wpis(ów)' % (k, len(p[k])))


if __name__ == '__main__':
    main()
