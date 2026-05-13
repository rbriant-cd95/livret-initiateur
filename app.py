"""
Livret Pédagogique Initiateur FFESSM — Application web
=======================================================
Déploiement : Streamlit Cloud (gratuit)
https://streamlit.io/cloud

Usage :
  1. L'utilisateur uploade le fichier Excel des séances
  2. L'utilisateur uploade le PDF du livret vierge
  3. Il clique sur "Générer le livret"
  4. Il télécharge le PDF rempli
"""

import io
import streamlit as st
import openpyxl
from datetime import datetime
from pypdf import PdfReader, PdfWriter
from pypdf.annotations import FreeText
from pypdf.generic import NameObject, NumberObject


# ============================================================
#  PARAMÈTRES (modifiables par le déployeur)
# ============================================================
SCALE = 1.2   # Facteur d'échelle des polices (1.0 = origine, 1.2 = +20%)
# ============================================================


# ------ Constantes PDF -------------------------------------
PDF_H            = 842.00
IMG_W            = 707
IMG_H            = 1000
PX_PER_PT        = IMG_H / PDF_H
CHAR_WIDTH       = 0.90
CHAR_PADDING     = 14
Q_MAP            = {'left': 0, 'center': 1}


# ------ Utilitaires ----------------------------------------

def fmt(d):
    if isinstance(d, datetime): return d.strftime('%d/%m/%y')
    return str(d).strip() if d else ''

def s(v): return str(v).strip() if v else ''

def fs(base): return int(base * SCALE)

def clean_theme(t):
    if not t: return ''
    t = str(t).strip()
    for p in ['Codep  ', 'Codep - ', 'Codep -', 'Codep ']:
        if t.startswith(p): return t[len(p):]
    return t

def vcenter(cell_y0, cell_y1, font_size_pts, n_lines=1, line_spacing=1.35):
    text_h_px   = font_size_pts * PX_PER_PT * line_spacing * n_lines
    cell_center = (cell_y0 + cell_y1) / 2
    y0 = max(cell_center - text_h_px / 2, cell_y0 + 2)
    y1 = min(cell_center + text_h_px / 2, cell_y1 - 2)
    return int(y0), int(y1)

def hcenter_eval(cell_x0, cell_x1, text, font_size_pts):
    text_w  = len(text) * font_size_pts * CHAR_WIDTH + CHAR_PADDING
    cell_cx = (cell_x0 + cell_x1) / 2
    x0 = max(cell_cx - text_w / 2, cell_x0 + 3)
    x1 = min(cell_cx + text_w / 2, cell_x1 - 3)
    return int(x0), int(x1)

def estimate_lines(text, font_size_pts, cell_width_px):
    if not text: return 1
    chars_per_line = cell_width_px / (font_size_pts * 0.55 * PX_PER_PT)
    return max(1, round(len(text) / max(chars_per_line, 1) + 0.5))


# ------ Lecture Excel --------------------------------------

def read_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    def row(i):
        return list(ws.iter_rows(min_row=i, max_row=i, values_only=True))[0]

    m1 = []
    for i in range(3, 44):
        r = row(i)
        if r[0] is None: continue
        m1.append({
            'date': fmt(r[0]), 'niveau': s(r[1]), 'theme': clean_theme(r[2]),
            'formateur': s(r[3]) or s(r[14]),
            'objectifs':   s(r[4]),  'positionner': s(r[5]),
            'justifier':   s(r[6]),  'strategie':   s(r[7]),
            'accueil':     s(r[8]),  'animer':      s(r[9]),
            'mettre':      s(r[10]), 'evaluer':     s(r[11]),
        })

    m2 = []
    for i in range(50, 56):
        r = row(i)
        if all(v is None for v in r): continue
        m2.append({
            'date': fmt(r[0]), 'theme': s(r[2]),
            'formateur': s(r[3]) or s(r[10]),
            'accueil': s(r[4]), 'organiser': s(r[5]),
            'securiser': s(r[6]), 'reagir': s(r[7]),
        })

    m3 = []
    for i in range(59, 61):
        r = row(i)
        m3.append({
            'date': fmt(r[0]), 'theme': s(r[2]),
            'formateur': s(r[3]) or s(r[10]),
            'identifier': s(r[4]), 'planifier': s(r[5]),
            'logistique': s(r[6]), 'moyens': s(r[7]),
        })

    return m1, m2, m3


# ------ Construction des champs ----------------------------

def build_fields(m1, m2, m3):
    fields = []
    lc = [0]

    def add(pg, desc, bbox, text, font_size, h_align='left'):
        if not text: return
        lc[0] += 1
        lx = 40 + (lc[0] % 5) * 2
        fields.append({
            "page_number": pg, "description": desc, "field_label": desc,
            "label_bounding_box": [lx, 10 + lc[0], lx + 1, 11 + lc[0]],
            "entry_bounding_box": bbox,
            "entry_text": {"text": str(text), "font_size": font_size, "h_align": h_align},
        })

    def field(pg, desc, x0, x1, cy0, cy1, text, font_size, n_lines=None):
        if not text: return
        nl = n_lines or estimate_lines(text, font_size, x1 - x0)
        y0, y1 = vcenter(cy0, cy1, font_size, nl)
        add(pg, desc, [x0 + 3, y0, x1 - 3, y1], text, font_size)

    def field_eval(pg, desc, cx0, cx1, cy0, cy1, text, font_size):
        if not text: return
        bx0, bx1 = hcenter_eval(cx0, cx1, text, font_size)
        by0, by1 = vcenter(cy0, cy1, font_size, n_lines=1)
        add(pg, desc, [bx0, by0, bx1, by1], text, font_size, h_align='center')

    # Module 1
    M1C = [(207,357),(357,506),(506,656)]
    M1R = {
        'date':(322,359),'formateur':(359,398),'niveau':(398,445),'theme':(445,500),
        'objectifs':(502,541),'positionner':(541,596),'justifier':(596,634),
        'strategie':(634,684),'accueil':(684,722),'animer':(722,759),
        'mettre':(759,815),'evaluer':(817,870),
    }
    EVAL1 = {'objectifs','positionner','justifier','strategie','accueil','animer','mettre','evaluer'}

    for si, sess in enumerate(m1):
        page = 5 + si // 3; ci = si % 3
        x0, x1 = M1C[ci]; p = f"p{page}_c{ci+1}"
        field(page,f"{p}_dt",x0,x1,*M1R['date'],      sess['date'],      fs(8), n_lines=1)
        field(page,f"{p}_fo",x0,x1,*M1R['formateur'], sess['formateur'], fs(7))
        field(page,f"{p}_ni",x0,x1,*M1R['niveau'],    sess['niveau'],    fs(8), n_lines=1)
        t = sess['theme']
        fs_t = fs(5) if len(t)>50 else (fs(6) if len(t)>35 else fs(7))
        field(page,f"{p}_th",x0,x1,*M1R['theme'],t,fs_t)
        for rk in EVAL1:
            v = sess.get(rk,'')
            if v: field_eval(page,f"{p}_{rk[:3]}",x0,x1,*M1R[rk],v,fs(9))

    # Module 2
    M2C = [(151,319),(319,488),(488,656)]
    M2R = {
        'date':(340,373),'formateur':(373,425),'theme':(425,478),
        'accueil':(478,529),'organiser':(529,588),'securiser':(588,640),'reagir':(640,693),
    }
    EVAL2 = {'accueil','organiser','securiser','reagir'}

    for si, sess in enumerate(m2):
        page = 25 + si // 3; ci = si % 3
        x0, x1 = M2C[ci]; p = f"p{page}_c{ci+1}"
        field(page,f"{p}_dt",x0,x1,*M2R['date'],      sess['date'],      fs(8), n_lines=1)
        field(page,f"{p}_fo",x0,x1,*M2R['formateur'], sess['formateur'], fs(7))
        t = sess['theme']
        fs_t = fs(5) if len(t)>50 else (fs(6) if len(t)>30 else fs(7))
        field(page,f"{p}_th",x0,x1,*M2R['theme'],t,fs_t)
        for rk in EVAL2:
            v = sess.get(rk,'')
            if v: field_eval(page,f"{p}_{rk[:3]}",x0,x1,*M2R[rk],v,fs(9))

    # Module 3
    M3C = [(151,319),(319,488),(488,656)]
    M3S = [
        {'date':(355,375),'formateur':(375,400),'niveau':(400,441),
         'identifier':(441,502),'planifier':(502,543),'logistique':(543,590),'moyens':(590,637)},
        {'date':(653,674),'formateur':(674,714),'niveau':(714,755),
         'identifier':(755,816),'planifier':(816,857),'logistique':(857,904),'moyens':(904,951)},
    ]
    EVAL3 = {'identifier','planifier','logistique','moyens'}

    for si, sub in enumerate([0,1]):
        if si >= len(m3): break
        sess = m3[si]; ci = 0
        x0, x1 = M3C[ci]; p = f"p29_st{sub+1}_c1"
        field(29,f"{p}_dt",x0,x1,*M3S[sub]['date'],      sess['date'],      fs(8), n_lines=1)
        field(29,f"{p}_fo",x0,x1,*M3S[sub]['formateur'], sess['formateur'], fs(7))
        t = sess['theme']
        fs_t = fs(5) if len(t)>50 else (fs(6) if len(t)>30 else fs(7))
        field(29,f"{p}_th",x0,x1,*M3S[sub]['niveau'],t,fs_t)
        for rk in EVAL3:
            v = sess.get(rk,'')
            if v: field_eval(29,f"{p}_{rk[:3]}",x0,x1,*M3S[sub][rk],v,fs(9))

    pages_used = sorted(set(f['page_number'] for f in fields))
    return {
        "pages": [{"page_number": p, "image_width": IMG_W, "image_height": IMG_H}
                  for p in pages_used],
        "form_fields": fields,
    }


# ------ Remplissage PDF ------------------------------------

def fill_pdf(pdf_bytes, fields_data):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    writer.append(reader)

    pdf_dims = {i+1: [float(p.mediabox.width), float(p.mediabox.height)]
                for i, p in enumerate(reader.pages)}
    count = 0

    for f in fields_data["form_fields"]:
        pg = f["page_number"]
        pi = next(p for p in fields_data["pages"] if p["page_number"] == pg)
        pdf_w, pdf_h = pdf_dims[pg]
        iw, ih = pi["image_width"], pi["image_height"]
        bb = f["entry_bounding_box"]

        xs = pdf_w / iw; ys = pdf_h / ih
        left   = bb[0] * xs; right  = bb[2] * xs
        top    = pdf_h - bb[1] * ys; bottom = pdf_h - bb[3] * ys

        et = f.get("entry_text", {})
        text = et.get("text", "")
        if not text: continue

        ann = FreeText(
            text=text, rect=(left, bottom, right, top),
            font=et.get("font", "Arial"),
            font_size=f"{et.get('font_size', 8)}pt",
            font_color=et.get("font_color", "000000"),
            border_color=None, background_color=None,
        )
        ann[NameObject("/Q")] = NumberObject(Q_MAP.get(et.get("h_align","left"), 0))
        ann[NameObject("/F")] = NumberObject(4)  # flag Print
        writer.add_annotation(page_number=pg - 1, annotation=ann)
        count += 1

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.getvalue(), count


# ============================================================
#  INTERFACE STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Livret Initiateur FFESSM",
    page_icon="🤿",
    layout="centered",
)

st.title("🤿 Livret Pédagogique Initiateur FFESSM")
st.markdown("Remplissage automatique du livret à partir du fichier Excel de séances.")

st.divider()

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader(
        "📊 Fichier Excel des séances",
        type=["xlsx"],
        help="Seances_initiateur.xlsx",
    )
with col2:
    pdf_file = st.file_uploader(
        "📄 Livret PDF vierge",
        type=["pdf"],
        help="Livret_pédagogique_Initiateur__1_.pdf",
    )

st.divider()

if excel_file and pdf_file:
    if st.button("⚙️ Générer le livret rempli", type="primary", use_container_width=True):
        with st.spinner("Traitement en cours…"):
            try:
                m1, m2, m3 = read_excel(excel_file.read())
                fields_data = build_fields(m1, m2, m3)
                pdf_out, count = fill_pdf(pdf_file.read(), fields_data)

                st.success(
                    f"✅ Livret généré — {count} annotations  "
                    f"({len(m1)} séances M1, {len(m2)} M2, {len(m3)} M3)"
                )
                st.download_button(
                    label="⬇️ Télécharger le livret rempli",
                    data=pdf_out,
                    file_name="Livret_initiateur_rempli.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    type="primary",
                )
            except Exception as e:
                st.error(f"❌ Erreur : {e}")
                st.exception(e)
else:
    st.info("👆 Uploadez les deux fichiers pour activer la génération.")

st.divider()
st.caption("FFESSM Codep 95 — Script généré avec Claude (Anthropic)")
