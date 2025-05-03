import os
import re
import locale
import logging
import requests
from datetime import datetime
from isbnlib import meta
from habanero import Crossref
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# ==== CONFIGURAÇÕES GERAIS ====
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    print("⚠️ Locale pt_BR.UTF-8 não disponível. Usando o padrão.")

logging.basicConfig(
    filename='error_log.txt',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def log_error(msg: str):
    logging.error(msg)

def cm_to_pt(cm: float) -> Pt:
    return Pt(cm * 28.35)

def set_document_styles(doc: Document, lang: str):
    sec = doc.sections[0]
    sec.top_margin = cm_to_pt(3)
    sec.right_margin = cm_to_pt(3)
    sec.bottom_margin = cm_to_pt(2)
    sec.left_margin = cm_to_pt(2)
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(12)
    style.paragraph_format.line_spacing = 1.0
    for para in doc.paragraphs:
        for run in para.runs:
            rPr = run._element.get_or_add_rPr()
            rPr.set(qn('w:lang'), lang)

# ==== PADRÕES DE IDENTIFICAÇÃO ====
DOI_RE  = re.compile(r'^(?:https?://doi\.org/)?10\.\d{4,9}/\S+$')
ISBN_RE = re.compile(r'^(97(8|9))?\d{9}(\d|X)$', re.IGNORECASE)

# ==== OVERRIDES PARA DOIs ESPECÍFICOS (Zenodo, etc.) ====
ZENODO_OVERRIDES = {
    '10.5281/zenodo.5829447': {
        'container': 'IA Policy Brief Series',
        'volume': '1',
        'number': '1',
        'pages': '1–3'
    }
}

# ==== NORMALIZAÇÃO DE TÍTULOS ====
def normalize_title(title: str) -> str:
    """
    Se o título vier todo em MAIÚSCULAS, converte para Title Case.
    Caso contrário, retorna o título original.
    """
    letters = re.sub(r'[^A-Za-zÀ-ÖØ-öø-ÿ]+', '', title)
    if letters and letters == letters.upper():
        return title.title()
    return title

# ==== EXTRAÇÃO MANUAL ====
def extract_reference_parts(reference: str):
    m = re.match(r'^(.*?)\.\s*', reference)
    authors = m.group(1).strip() if m else ''
    ym = re.search(r'\b(\d{4})\b', reference)
    year = ym.group(1) if ym else ''
    tm = re.search(r'\.\s*(.*?)\.\s*', reference)
    title = tm.group(1).strip() if tm else ''
    after = reference.split(f'. {title}.')[-1]
    after = re.sub(r'\b' + re.escape(year) + r'\b', '', after).strip(' ,.')
    if ':' in after:
        city, publisher = map(str.strip, after.split(':', 1))
    elif ',' in after:
        city, publisher = map(str.strip, after.split(',', 1))
    else:
        city, publisher = '', after.strip()
    return authors, year, title, city, publisher

# ==== FORMATAÇÃO DE AUTORES ====
def format_author_abnt(authors: str) -> str:
    out = []
    for name in authors.split(';'):
        parts = re.sub(r'[^\w\s]', '', name.strip()).split()
        if len(parts) >= 2:
            last = parts[-1].upper()
            firsts = ' '.join(p.title() for p in parts[:-1])
            out.append(f"{last}, {firsts}")
        else:
            out.append(parts[0].upper())
    return '; '.join(out) + '.'

def format_author_apa7(authors: str) -> str:
    out = []
    for name in authors.split(';'):
        n = name.strip()
        if ',' in n:
            last, given = map(str.strip, n.split(',', 1))
            parts = given.split()
        else:
            allp = n.split()
            last = allp[-1]
            parts = allp[:-1]
        initials = ' '.join(f"{p[0].upper()}." for p in parts if p)
        out.append(f"{last.capitalize()}, {initials}")
    return '; '.join(out)

# ==== CONSULTA EXTERNA (DOI / ISBN) ====
def get_data_by_doi(doi: str):
    key = doi.split('https://doi.org/')[-1]
    cr = Crossref()
    try:
        return cr.works(ids=key)['message']
    except Exception as e:
        log_error(f"Crossref fetch failed for DOI {key}: {e}")
    try:
        resp = requests.get(f"https://api.datacite.org/dois/{key}")
        resp.raise_for_status()
        return resp.json().get('data', {}).get('attributes')
    except Exception as e:
        log_error(f"DataCite fetch failed for DOI {key}: {e}")
    return None

import re
import requests

def get_data_by_isbn(isbn: str):
    """
    Tenta primeiro com isbnlib.meta; se falhar ou não trouxer Year, faz fallback
    na OpenLibrary, extraindo título, autores, editora e ano.
    Retorna dict com chaves:
      - 'Title': str
      - 'Authors': [str, …]
      - 'Publisher': str
      - 'Year': str
    ou None se nada foi obtido.
    """
    key = isbn.replace('-', '').strip()
    raw = None

    # 1) tentativa com isbnlib.meta
    try:
        raw = meta(key)
    except Exception as e:
        log_error(f"isbnlib.meta failed for ISBN {isbn}: {e}")
        raw = None

    # se meta() trouxe Title e Year, devolvemos imediatamente
    if raw and raw.get('Title') and raw.get('Year'):
        return raw

    # Extrai o que vier de meta(), mesmo que incompleto
    title     = raw.get('Title', '')     if raw else ''
    authors   = raw.get('Authors', [])   if raw else []
    publisher = raw.get('Publisher', '') if raw else ''
    year_meta = raw.get('Year', '')      if raw else ''

    # 2) fallback OpenLibrary
    try:
        url = f"https://openlibrary.org/api/books?bibkeys=ISBN:{key}&format=json&jscmd=data"
        resp = requests.get(url)
        resp.raise_for_status()
        record = resp.json().get(f"ISBN:{key}")
        if record:
            # extrai ano de publish_date
            pub_date   = record.get('publish_date', '')
            m_year     = re.search(r'\b(\d{4})\b', pub_date)
            year_ol    = m_year.group(1) if m_year else year_meta

            return {
                'Title':     title or record.get('title', ''),
                'Authors':   authors or [a.get('name','') for a in record.get('authors', [])],
                'Publisher': publisher or ', '.join(p.get('name','') for p in record.get('publishers', [])),
                'Year':      year_ol or ''
            }
    except Exception as e:
        log_error(f"OpenLibrary fallback failed for ISBN {isbn}: {e}")

    # 3) se meta() trouxe algo, mesmo sem Year, devolve com Year possivelmente vazio
    if raw:
        return {
            'Title':     title,
            'Authors':   authors,
            'Publisher': publisher,
            'Year':      year_meta or ''
        }

    # nada encontrado
    return None

# ==== GERAÇÃO DE PARTES DE REFERÊNCIA ====
def process_reference_abnt(
    authors: str, title: str,
    city: str, publisher: str, year: str,
    additional_info: str = None,
    volume: str = None, number: str = None, pages: str = None,
    doi: str = None, url: str = None, container: str = None
) -> list:
    parts = [f"{authors} "]
    if container:
        parts.append(f"{title}. ")
        parts.append({"text": container, "bold": True}); parts.append(". ")
    else:
        parts.append({"text": title, "bold": True}); parts.append(". ")
    if not container and publisher:
        if city:
            parts.append(f"{city}: {publisher}, {year}. ")
        else:
            parts.append(f"{publisher}, {year}. ")
    if additional_info: parts.append(f"{additional_info} ")
    if volume: parts.append(f"v. {volume} ")
    if number: parts.append(f"n. {number} ")
    if pages: parts.append(f"p. {pages}. ")
    if doi: parts.append(f"DOI: {doi}. ")
    if url: parts.append(f"Disponível em: {url}. ")
    return parts

def process_reference_apa7(
    authors: str, title: str, year: str,
    publisher: str = None, volume: str = None, number: str = None,
    pages: str = None, doi: str = None, url: str = None, container: str = None
) -> list:
    parts = [f"{authors} ({year}). "]
    if container:
        parts.append(f"{title}. ")
        parts.append({"text": container, "bold": True}); parts.append(". ")
    else:
        parts.append({"text": title, "bold": True}); parts.append(". ")
    if volume and number and pages:
        parts.append(f"{volume}({number}), {pages}. ")
    elif pages:
        parts.append(f"{pages}. ")
    if not container and publisher:
        parts.append(f"{publisher}. ")
    if doi:
        parts.append(f"https://doi.org/{doi}")
    elif url:
        parts.append(f"Disponível em: {url}")
    return parts

# ==== FLUXO PRINCIPAL ====
def process_reference(ref: str, ref_type: str = "manual") -> (list, list):
    try:
        if ref_type == "doi":
            data = get_data_by_doi(ref)
            if not data:
                raise ValueError(f"nenhum dado DOI para {ref}")

            # Crossref vs DataCite
            if 'container-title' in data:
                raw_authors = '; '.join(f"{a['given']} {a['family']}" for a in data.get('author', []))
                title     = normalize_title(data.get('title', [''])[0])
                container = data.get('container-title', [''])[0] or None
                year      = str(data.get('issued', {}).get('date-parts', [[None]])[0][0] or '')
                volume    = data.get('volume')
                number    = data.get('issue')
                pages     = data.get('page')
                url       = data.get('URL')
            elif 'titles' in data:
                raw_authors = '; '.join(
                    f"{c.get('givenName','')} {c.get('familyName','')}".strip()
                    for c in data.get('creators', [])
                )
                title     = normalize_title(data.get('titles', [{}])[0].get('title',''))
                container = data.get('publisher')
                year      = str(data.get('publicationYear',''))
                volume    = data.get('volume')
                number    = data.get('issue')
                pages     = data.get('page')
                url       = data.get('url')
            else:
                raise ValueError(f"Formato inesperado de metadados DOI para {ref}")

            doi_key = ref.split('https://doi.org/')[-1]

            # override Zenodo específico
            override = ZENODO_OVERRIDES.get(doi_key, {})
            container = override.get('container', container)
            volume    = override.get('volume',    volume)
            number    = override.get('number',    number)
            pages     = override.get('pages',     pages)

            abnt = process_reference_abnt(
                format_author_abnt(raw_authors), title,
                city='', publisher='', year=year,
                volume=volume, number=number, pages=pages,
                doi=doi_key, url=url, container=container
            )
            apa7 = process_reference_apa7(
                format_author_apa7(raw_authors), title, year,
                publisher=None, volume=volume, number=number,
                pages=pages, doi=doi_key, url=url, container=container
            )

        elif ref_type == "isbn":
            data = get_data_by_isbn(ref)
            if not data:
                raise ValueError(f"nenhum dado ISBN para {ref}")
            raw_authors = '; '.join(data.get('Authors', []))
            title       = normalize_title(data.get('Title',''))
            year        = data.get('Year','')
            publisher   = data.get('Publisher','')
            abnt = process_reference_abnt(
                format_author_abnt(raw_authors), title,
                city='', publisher=publisher, year=year
            )
            apa7 = process_reference_apa7(
                format_author_apa7(raw_authors), title, year, publisher=publisher
            )

        else:
            raw_authors, year, title, city, publisher = extract_reference_parts(ref)
            title = normalize_title(title)
            abnt = process_reference_abnt(
                format_author_abnt(raw_authors), title,
                city=city, publisher=publisher, year=year
            )
            apa7 = process_reference_apa7(
                format_author_apa7(raw_authors), title, year, publisher=publisher
            )

        return abnt, apa7

    except Exception as e:
        log_error(f"Erro processando '{ref}': {e}")
        return [], []

def add_formatted_reference(paragraph, parts: list):
    for part in parts:
        if isinstance(part, dict) and part.get("bold"):
            paragraph.add_run(part["text"]).bold = True
        else:
            paragraph.add_run(part if isinstance(part, str) else part.get("text",""))

def save_references(refs: list, filename: str, lang: str, heading: str):
    doc = Document()
    set_document_styles(doc, lang)
    doc.add_heading(heading, level=1)
    for parts in refs:
        p = doc.add_paragraph()
        p.alignment = 3
        add_formatted_reference(p, parts)
    doc.save(filename)

def process_references_from_file(input_file: str):
    if not os.path.exists(input_file):
        print(f"Arquivo não encontrado: {input_file}")
        return
    abnt_list, apa7_list = [], []
    with open(input_file, encoding='utf-8') as f:
        for line in f:
            ref = line.strip()
            if not ref:
                continue
            if DOI_RE.match(ref):
                abnt, apa7 = process_reference(ref, "doi")
            elif ISBN_RE.match(ref.replace('-', '')):
                abnt, apa7 = process_reference(ref, "isbn")
            else:
                abnt, apa7 = process_reference(ref, "manual")
            if abnt and apa7:
                abnt_list.append(abnt)
                apa7_list.append(apa7)
    base = os.path.dirname(os.path.abspath(input_file))
    save_references(abnt_list, os.path.join(base, 'referencias_abnt.docx'), 'pt-BR', 'REFERÊNCIAS')
    save_references(apa7_list, os.path.join(base, 'referencias_apa7.docx'), 'en-US', 'REFERENCES')

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    process_references_from_file(os.path.join(script_dir, 'referencias.txt'))
