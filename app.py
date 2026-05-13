#!/usr/bin/env python3
"""
Remplissage automatique du Livret Pédagogique Initiateur FFESSM
================================================================
Pré-requis :
    pip install openpyxl pypdf

Usage :
    python remplir_livret_initiateur_v4.py

Adapter les chemins et le facteur d'échelle (section CONFIGURATION).

Centrage des codes A / ECA / NT :
  - Horizontal : centrage géométrique (bbox de la largeur exacte du texte,
                 positionnée au centre de la colonne). Indépendant du lecteur PDF.
  - Vertical   : bbox réduite à la hauteur d'une ligne, positionnée au centre
                 de la cellule.
Les autres champs (date, formateur, thème, niveau) ne sont pas modifiés.

Structure Excel attendue :
  Feuil1
  - Lignes 3-43  : Module 1 — Enseignement pratique
      Col A=Date, B=Niveau, C=Thème, D=Formateur, E=Objectifs, F=Positionnement,
      G=Justification, H=Stratégie, I=Accueil, J=Animer, K=Mettre en œuvre,
      L=Réaliser éval., O=Évaluateur(s)
  - Lignes 50-55 : Module 2 — Organiser l'activité
      Col A=Date, C=Thème, D=Formateur, E=Accueillir, F=Organiser,
      G=Sécuriser, H=Réagir, K=Évaluateur(s)
  - Lignes 59-60 : Module 3 — Organiser un cursus
      Col A=Date, C=Thème, D=Formateur, E=Identifier, F=Planifier,
      G=Logistique, H=Moyens, K=Évaluateur(s)

Structure PDF attendue (35 pages) :
  Pages  5-23 : tableaux Enseignement pratique (3 séances / page)
  Pages 25-27 : tableaux Organiser l'activité  (3 séances / page)
  Page     29 : tableau Organiser un cursus   (2 sous-tableaux)
"""

import openpyxl
from datetime import datetime
from pypdf import PdfReader, PdfWriter
from pypdf.annotations import FreeText
from pypdf.generic import NameObject, NumberObject


# ============================================================
#  CONFIGURATION — à adapter selon vos chemins de fichiers
# ============================================================
EXCEL_PATH = "Seances_initiateur.xlsx"
PDF_INPUT  = "Livret_pédagogique_Initiateur__1_.pdf"
PDF_OUTPUT = "Livret_initiateur_rempli.pdf"

# Facteur d'échelle pour la taille des polices.
# 1.0 = taille d'origine, 1.2 = +20%, 1.4 = +40%
# Au-delà de 1.4, risque de débordement sur les noms/thèmes longs.
SCALE = 1.3
# ============================================================


# ------ Constantes PDF -------------------------------------
# Dimensions de la page A4 en points (vérifié sur ce PDF)
PDF_H = 842.00
# Dimensions de l'image de référence utilisée pour les coordonnées
IMG_W = 707
IMG_H = 1000
# Rapport px → pt (pour convertir une taille de police pts en pixels image)
PX_PER_PT = IMG_H / PDF_H   # ≈ 1.187 px/pt

# Largeur d'un caractère Arial en px/pt (mesurée empiriquement sur ce PDF)
# On utilise un coefficient large (0.90) pour éviter tout débordement,
# plus une marge fixe pour les paddings internes de l'annotation FreeText.
CHAR_WIDTH_PX_PER_PT = 0.90
CHAR_PADDING_PX      = 14    # marge interne FreeText (~6pt chaque côté)

# Correspondance h_align → /Q PDF (0=left, 1=center, 2=right)
Q_MAP = {'left': 0, 'center': 1, 'right': 2}


# ------ Utilitaires ----------------------------------------

def fmt(d):
    if isinstance(d, datetime): return d.strftime('%d/%m/%y')
    return str(d).strip() if d else ''

def s(v): return str(v).strip() if v else ''

def fs(base):
    """Applique le facteur SCALE à une taille de police de base."""
    return int(base * SCALE)

def clean_theme(t):
    """Supprime le préfixe 'Codep ...' des thèmes du module 1."""
    if not t: return ''
    t = str(t).strip()
    for p in ['Codep  ', 'Codep - ', 'Codep -', 'Codep ']:
        if t.startswith(p): return t[len(p):]
    return t

def vcenter(cell_y0, cell_y1, font_size_pts, n_lines=1, line_spacing=1.35):
    """
    Calcule y0/y1 (coordonnées image) pour centrer verticalement n_lines
    lignes de texte dans la cellule [cell_y0, cell_y1].
    """
    text_h_px = font_size_pts * PX_PER_PT * line_spacing * n_lines
    cell_center = (cell_y0 + cell_y1) / 2
    y0 = cell_center - text_h_px / 2
    y1 = cell_center + text_h_px / 2
    y0 = max(y0, cell_y0 + 2)
    y1 = min(y1, cell_y1 - 2)
    return int(y0), int(y1)

def hcenter_eval(cell_x0, cell_x1, text, font_size_pts):
    """
    Calcule x0/x1 (coordonnées image) pour centrer géométriquement
    un code A/ECA/NT dans la colonne [cell_x0, cell_x1].
    La largeur du texte est estimée à partir de la taille de police.
    """
    text_w_px = len(text) * font_size_pts * CHAR_WIDTH_PX_PER_PT + CHAR_PADDING_PX
    cell_cx   = (cell_x0 + cell_x1) / 2
    x0 = cell_cx - text_w_px / 2
    x1 = cell_cx + text_w_px / 2
    # Clamp dans les limites de la colonne
    x0 = max(x0, cell_x0 + 3)
    x1 = min(x1, cell_x1 - 3)
    return int(x0), int(x1)

def estimate_lines(text, font_size_pts, cell_width_px, chars_per_pt=0.55):
    """Estime le nombre de lignes qu'occupera le texte dans la cellule."""
    if not text: return 1
    chars_per_line = cell_width_px / (font_size_pts * chars_per_pt * PX_PER_PT)
    return max(1, round(len(text) / max(chars_per_line, 1) + 0.5))


# ------ Lecture Excel --------------------------------------

def read_excel(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    def row(i):
        return list(ws.iter_rows(min_row=i, max_row=i, values_only=True))[0]

    # Module 1 — lignes 3 à 43
    m1 = []
    for i in range(3, 44):
        r = row(i)
        if r[0] is None: continue
        formateur = s(r[3]) or s(r[14])
        m1.append({
            'date': fmt(r[0]), 'niveau': s(r[1]), 'theme': clean_theme(r[2]),
            'formateur': formateur,
            'objectifs':   s(r[4]),  'positionner': s(r[5]),
            'justifier':   s(r[6]),  'strategie':   s(r[7]),
            'accueil':     s(r[8]),  'animer':      s(r[9]),
            'mettre':      s(r[10]), 'evaluer':     s(r[11]),
        })

    # Module 2 — lignes 50 à 55
    m2 = []
    for i in range(50, 56):
        r = row(i)
        if all(v is None for v in r): continue
        formateur = s(r[3]) or s(r[10])
        m2.append({
            'date': fmt(r[0]), 'theme': s(r[2]), 'formateur': formateur,
            'accueil':   s(r[4]), 'organiser': s(r[5]),
            'securiser': s(r[6]), 'reagir':    s(r[7]),
        })

    # Module 3 — lignes 59 à 60
    m3 = []
    for i in range(59, 61):
        r = row(i)
        formateur = s(r[3]) or s(r[10])
        m3.append({
            'date': fmt(r[0]), 'theme': s(r[2]), 'formateur': formateur,
            'identifier': s(r[4]), 'planifier':  s(r[5]),
            'logistique': s(r[6]), 'moyens':     s(r[7]),
        })

    return m1, m2, m3


# ------ Construction des champs ----------------------------

def build_fields(m1, m2, m3):
    fields = []
    lc = [0]

    def add(pg, desc, bbox, text, font_size, h_align='left'):
        """Ajoute un champ texte avec la bbox fournie."""
        if not text: return
        lc[0] += 1
        lx_label = 40 + (lc[0] % 5) * 2
        fields.append({
            "page_number": pg,
            "description": desc,
            "field_label": desc,
            "label_bounding_box": [lx_label, 10 + lc[0], lx_label + 1, 11 + lc[0]],
            "entry_bounding_box": bbox,
            "entry_text": {"text": str(text), "font_size": font_size, "h_align": h_align},
        })

    def field(pg, desc, x0, x1, cell_y0, cell_y1, text, font_size, n_lines=None):
        """Champ standard : aligné à gauche, centré verticalement."""
        if not text: return
        cell_w = x1 - x0
        nl = n_lines if n_lines is not None else estimate_lines(text, font_size, cell_w)
        y0, y1 = vcenter(cell_y0, cell_y1, font_size, nl)
        add(pg, desc, [x0 + 3, y0, x1 - 3, y1], text, font_size)

    def field_eval(pg, desc, col_x0, col_x1, cell_y0, cell_y1, text, font_size):
        """
        Champ évaluation (A / ECA / NT) : centrage géométrique H+V.
        - La bbox est centrée dans la colonne (centrage indépendant du lecteur).
        - /Q=1 est également appliqué pour les lecteurs qui le supportent.
        La bbox est volontairement plus large que le texte pour éviter tout
        retour à la ligne (les annotations FreeText ont un padding interne).
        """
        if not text: return
        bx0, bx1 = hcenter_eval(col_x0, col_x1, text, font_size)
        by0, by1 = vcenter(cell_y0, cell_y1, font_size, n_lines=1)
        add(pg, desc, [bx0, by0, bx1, by1], text, font_size, h_align='center')

    # ---- Module 1 : pages 5-23, 3 séances par page --------
    M1C = [(207, 357), (357, 506), (506, 656)]
    M1R = {
        'date':        (322, 359), 'formateur':   (359, 398),
        'niveau':      (398, 445), 'theme':       (445, 500),
        'objectifs':   (502, 541), 'positionner': (541, 596),
        'justifier':   (596, 634), 'strategie':   (634, 684),
        'accueil':     (684, 722), 'animer':      (722, 759),
        'mettre':      (759, 815), 'evaluer':     (817, 870),
    }
    EVAL1 = {'objectifs', 'positionner', 'justifier', 'strategie',
             'accueil', 'animer', 'mettre', 'evaluer'}

    for si, sess in enumerate(m1):
        page = 5 + si // 3
        ci   = si % 3
        x0, x1 = M1C[ci]
        pref = f"p{page}_c{ci+1}"

        field(page, f"{pref}_dt", x0, x1, *M1R['date'],      sess['date'],      fs(8), n_lines=1)
        field(page, f"{pref}_fo", x0, x1, *M1R['formateur'], sess['formateur'], fs(7))
        field(page, f"{pref}_ni", x0, x1, *M1R['niveau'],    sess['niveau'],    fs(8), n_lines=1)
        t = sess['theme']
        fs_t = fs(5) if len(t) > 50 else (fs(6) if len(t) > 35 else fs(7))
        field(page, f"{pref}_th", x0, x1, *M1R['theme'], t, fs_t)

        for rk in EVAL1:
            v = sess.get(rk, '')
            if v:
                field_eval(page, f"{pref}_{rk[:3]}", x0, x1, *M1R[rk], v, fs(9))

    # ---- Module 2 : pages 25-27, 3 séances par page -------
    M2C = [(151, 319), (319, 488), (488, 656)]
    M2R = {
        'date':      (340, 373), 'formateur': (373, 425), 'theme':    (425, 478),
        'accueil':   (478, 529), 'organiser': (529, 588),
        'securiser': (588, 640), 'reagir':    (640, 693),
    }
    EVAL2 = {'accueil', 'organiser', 'securiser', 'reagir'}

    for si, sess in enumerate(m2):
        page = 25 + si // 3
        ci   = si % 3
        x0, x1 = M2C[ci]
        pref = f"p{page}_c{ci+1}"

        field(page, f"{pref}_dt", x0, x1, *M2R['date'],      sess['date'],      fs(8), n_lines=1)
        field(page, f"{pref}_fo", x0, x1, *M2R['formateur'], sess['formateur'], fs(7))
        t = sess['theme']
        fs_t = fs(5) if len(t) > 50 else (fs(6) if len(t) > 30 else fs(7))
        field(page, f"{pref}_th", x0, x1, *M2R['theme'], t, fs_t)

        for rk in EVAL2:
            v = sess.get(rk, '')
            if v:
                field_eval(page, f"{pref}_{rk[:3]}", x0, x1, *M2R[rk], v, fs(9))

    # ---- Module 3 : page 29, 2 sous-tableaux --------------
    M3C = [(151, 319), (319, 488), (488, 656)]
    M3S = [
        {
            'date':       (355, 375), 'formateur':  (375, 400), 'niveau':    (400, 441),
            'identifier': (441, 502), 'planifier':  (502, 543),
            'logistique': (543, 590), 'moyens':     (590, 637),
        },
        {
            'date':       (653, 674), 'formateur':  (674, 714), 'niveau':    (714, 755),
            'identifier': (755, 816), 'planifier':  (816, 857),
            'logistique': (857, 904), 'moyens':     (904, 951),
        },
    ]
    EVAL3 = {'identifier', 'planifier', 'logistique', 'moyens'}

    for si, sub in enumerate([0, 1]):
        if si >= len(m3): break
        sess = m3[si]
        ci   = 0
        x0, x1 = M3C[ci]
        pref = f"p29_st{sub+1}_c1"

        field(29, f"{pref}_dt", x0, x1, *M3S[sub]['date'],      sess['date'],      fs(8), n_lines=1)
        field(29, f"{pref}_fo", x0, x1, *M3S[sub]['formateur'], sess['formateur'], fs(7))
        t = sess['theme']
        fs_t = fs(5) if len(t) > 50 else (fs(6) if len(t) > 30 else fs(7))
        field(29, f"{pref}_th", x0, x1, *M3S[sub]['niveau'], t, fs_t)

        for rk in EVAL3:
            v = sess.get(rk, '')
            if v:
                field_eval(29, f"{pref}_{rk[:3]}", x0, x1, *M3S[sub][rk], v, fs(9))

    pages_used = sorted(set(f['page_number'] for f in fields))
    return {
        "pages": [{"page_number": p, "image_width": IMG_W, "image_height": IMG_H}
                  for p in pages_used],
        "form_fields": fields,
    }


# ------ Remplissage PDF ------------------------------------

def fill_pdf(input_path, fields_data, output_path):
    reader = PdfReader(input_path)
    writer = PdfWriter()
    writer.append(reader)

    pdf_dims = {i + 1: [float(p.mediabox.width), float(p.mediabox.height)]
                for i, p in enumerate(reader.pages)}

    count = 0
    for f in fields_data["form_fields"]:
        pg = f["page_number"]
        pi = next(p for p in fields_data["pages"] if p["page_number"] == pg)
        pdf_w, pdf_h = pdf_dims[pg]
        iw, ih = pi["image_width"], pi["image_height"]
        bb = f["entry_bounding_box"]

        # Conversion coordonnées image → PDF
        xs = pdf_w / iw;  ys = pdf_h / ih
        left   = bb[0] * xs
        right  = bb[2] * xs
        top    = pdf_h - bb[1] * ys
        bottom = pdf_h - bb[3] * ys

        et   = f.get("entry_text", {})
        text = et.get("text", "")
        if not text: continue

        ann = FreeText(
            text=text,
            rect=(left, bottom, right, top),
            font=et.get("font", "Arial"),
            font_size=f"{et.get('font_size', 8)}pt",
            font_color=et.get("font_color", "000000"),
            border_color=None,
            background_color=None,
        )
        # Centrage horizontal /Q (0=left, 1=center) — backup pour lecteurs compatibles
        h_align = et.get("h_align", "left")
        ann[NameObject("/Q")] = NumberObject(Q_MAP.get(h_align, 0))
        # Flag Print (bit 3 = 4) : indispensable pour que l'annotation s'imprime
        ann[NameObject("/F")] = NumberObject(4)
        writer.add_annotation(page_number=pg - 1, annotation=ann)
        count += 1

    with open(output_path, "wb") as out:
        writer.write(out)

    print(f"PDF rempli : {output_path}  ({count} annotations)")


# ------ Point d'entrée -------------------------------------

if __name__ == "__main__":
    print(f"Lecture Excel  : {EXCEL_PATH}")
    m1, m2, m3 = read_excel(EXCEL_PATH)
    print(f"  Module 1 : {len(m1)} séances")
    print(f"  Module 2 : {len(m2)} séances")
    print(f"  Module 3 : {len(m3)} séances")

    fields_data = build_fields(m1, m2, m3)
    print(f"Champs générés : {len(fields_data['form_fields'])}  (SCALE={SCALE})")

    print(f"Remplissage PDF : {PDF_INPUT} → {PDF_OUTPUT}")
    fill_pdf(PDF_INPUT, fields_data, PDF_OUTPUT)
