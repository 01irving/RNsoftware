# -*- coding: utf-8 -*-
"""Genera el PDF: Limpieza y esterilización del material de preparación
y administración (apartado 3.1.1), con imágenes ilustrativas de
Wikimedia Commons (licencias CC).
"""
import os

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
    Image, KeepTogether,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from PIL import Image as PILImage

# ── Fuentes (Windows) ──────────────────────────────────────────────
FDIR = r'C:\Windows\Fonts'
pdfmetrics.registerFont(TTFont('Times New Roman', os.path.join(FDIR, 'times.ttf')))
pdfmetrics.registerFont(TTFont('Times New Roman Bold', os.path.join(FDIR, 'timesbd.ttf')))
pdfmetrics.registerFont(TTFont('Times New Roman Italic', os.path.join(FDIR, 'timesi.ttf')))
pdfmetrics.registerFont(TTFont('Times New Roman BoldItalic', os.path.join(FDIR, 'timesbi.ttf')))
registerFontFamily(
    'Times New Roman',
    normal='Times New Roman',
    bold='Times New Roman Bold',
    italic='Times New Roman Italic',
    boldItalic='Times New Roman BoldItalic',
)

# ━━ Paleta de la serie (pdf.py palette.generate — misma familia que
#    «Fórmula, biberones y tetinas - ENeA SEA.pdf») ━━
# Intent: neutral | Mode: minimal | Harmony: split_complementary
ACCENT       = colors.HexColor('#23748e')
TEXT_PRIMARY = colors.HexColor('#212325')
TEXT_MUTED   = colors.HexColor('#777c83')
BG_SURFACE   = colors.HexColor('#d8dce1')
BG_PAGE      = colors.HexColor('#f1f2f3')

OUT_DIR = r'D:\Biblioteca\Nutrición pediatrica\RNsoftware\Fórmula pediatrica'
OUT = os.path.join(OUT_DIR,
                   'Limpieza y esterilización del material - Apartado 3.1.1.pdf')

doc = SimpleDocTemplate(
    OUT, pagesize=A4,
    leftMargin=2.0 * cm, rightMargin=2.0 * cm,
    topMargin=2.2 * cm, bottomMargin=2.2 * cm,
    title='Limpieza y esterilización del material de preparación y administración',
    author='Z.ai', creator='Z.ai',
    subject='Procedimiento de limpieza y esterilización del material para preparar '
            'y administrar las tomas (apartado 3.1.1)',
)

# ── Estilos ────────────────────────────────────────────────────────
title_style = ParagraphStyle(
    'TitleDoc', fontName='Times New Roman', fontSize=18, leading=24,
    alignment=TA_CENTER, textColor=TEXT_PRIMARY, spaceAfter=4,
)
subtitle_style = ParagraphStyle(
    'SubtitleDoc', fontName='Times New Roman', fontSize=11.5, leading=16,
    alignment=TA_CENTER, textColor=TEXT_MUTED, spaceAfter=12,
)
intro_style = ParagraphStyle(
    'IntroCallout', fontName='Times New Roman', fontSize=11, leading=16.5,
    alignment=TA_JUSTIFY, textColor=TEXT_PRIMARY,
)
h2_style = ParagraphStyle(
    'H2', fontName='Times New Roman', fontSize=13.5, leading=17,
    alignment=TA_LEFT, textColor=ACCENT, spaceBefore=16, spaceAfter=8,
)
step_style = ParagraphStyle(
    'Step', fontName='Times New Roman', fontSize=10.5, leading=15.5,
    alignment=TA_JUSTIFY, textColor=TEXT_PRIMARY,
    leftIndent=18, firstLineIndent=-18, spaceAfter=8,
)
substep_style = ParagraphStyle(
    'SubStep', fontName='Times New Roman', fontSize=10.5, leading=15.5,
    alignment=TA_JUSTIFY, textColor=TEXT_PRIMARY,
    leftIndent=38, firstLineIndent=-18, spaceAfter=7,
)
caption_style = ParagraphStyle(
    'Caption', fontName='Times New Roman', fontSize=8.5, leading=11.5,
    alignment=TA_CENTER, textColor=TEXT_MUTED, spaceBefore=5,
)
note_style = ParagraphStyle(
    'Note', fontName='Times New Roman', fontSize=9, leading=12.5,
    alignment=TA_LEFT, textColor=TEXT_MUTED, spaceBefore=12,
)

# ── Imágenes (aportadas por el usuario) ────────────────────────────
IMG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       'data', 'img_limpieza')


def fit_image(path, max_width, max_height):
    """Inserta la imagen escalada proporcionalmente dentro del recuadro."""
    pil_img = PILImage.open(path)
    w, h = pil_img.size
    ratio = min(max_width / w, max_height / h)
    img = Image(path, width=w * ratio, height=h * ratio)
    img.hAlign = 'CENTER'
    return img


def par_imagen(path, didactic_caption, max_w=220, max_h=150):
    """Imagen con su pie, centrado como bloque."""
    return [fit_image(path, max_w, max_h),
            Paragraph(didactic_caption, caption_style)]


img_manos = par_imagen(
    os.path.join(IMG_DIR, 'manos.jpg'),
    'Lavado de manos con agua y jabón (pasos 1 y 5)', max_w=150, max_h=150)
img_biberon = par_imagen(
    os.path.join(IMG_DIR, 'biberon.jpg'),
    'Limpieza del biberón con cepillo y agua jabonosa (paso 2)',
    max_w=150, max_h=150)
img_secado = par_imagen(
    os.path.join(IMG_DIR, 'secado.jpg'),
    'Secado del material tras el enjuague (paso 3)', max_w=150, max_h=150)
img_hervido = par_imagen(
    os.path.join(IMG_DIR, 'hervido.jpg'),
    'Esterilización por ebullición en recipiente con tapadera (paso 4)',
    max_w=150, max_h=150)
img_guardado = par_imagen(
    os.path.join(IMG_DIR, 'guardado.jpg'),
    'Guardado del material cubierto en un lugar limpio (paso 6)',
    max_w=150, max_h=150)

fila_imagenes = Table(
    [[img_manos, img_biberon, img_secado]],
    colWidths=[(A4[0] - doc.leftMargin - doc.rightMargin) / 3.0] * 3)
fila_imagenes.setStyle(TableStyle([
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
]))

fila_imagenes_2 = Table(
    [[img_hervido, img_guardado]],
    colWidths=[(A4[0] - doc.leftMargin - doc.rightMargin) / 2.0] * 2)
fila_imagenes_2.setStyle(TableStyle([
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
]))

# ── Contenido (apartado 3.1.1) ─────────────────────────────────────
intro = Paragraph(
    'Es sumamente importante que todo el material empleado para alimentar a '
    'los lactantes y para preparar las tomas haya sido concienzudamente '
    'limpiado y esterilizado antes de usarlo.',
    intro_style,
)
intro_callout = Table([[intro]], colWidths=[A4[0] - doc.leftMargin - doc.rightMargin])
intro_callout.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), BG_PAGE),
    ('LINEBEFORE', (0, 0), (0, -1), 3, ACCENT),
    ('LEFTPADDING', (0, 0), (-1, -1), 12),
    ('RIGHTPADDING', (0, 0), (-1, -1), 12),
    ('TOPPADDING', (0, 0), (-1, -1), 9),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 9),
]))

pasos_1_3 = [
    ('1.', 'Siempre se lavarán las manos debidamente con agua y jabón antes '
           'de limpiar y esterilizar todo el material empleado en la '
           'preparación de las tomas.'),
    ('2.', '<b>Limpieza:</b> se lavará concienzudamente todo el material de '
           'preparación y administración en agua jabonosa caliente. Cuando '
           'se utilicen biberones, se utilizarán cepillos especiales limpios '
           'para frotar el interior y el exterior de los biberones y tetinas '
           'a fin de eliminar los restos de la toma anterior.'),
    ('3.', 'Después de lavar el material, se enjuagará debidamente con agua '
           'limpia.'),
]
paso_4 = ('4.', '<b>Esterilización:</b> si se utiliza un esterilizador '
                'comercial (esterilizador a vapor eléctrico o de microondas, '
                'o esterilizador químico), se seguirán las instrucciones del '
                'fabricante. Si se esteriliza el material por ebullición:')
subpasos = [
    ('a.', 'Se llenará un recipiente grande con agua y se sumergirá por '
           'completo todo el material lavado, velando por que no queden '
           'burbujas de aire atrapadas en el interior.'),
    ('b.', 'Se cubrirá el recipiente con una tapadera y se dejará hervir '
           'fuertemente el contenido, cuidando de que el agua no llegue a '
           'evaporarse por completo.'),
    ('c.', 'Se mantendrá el recipiente cubierto hasta que se necesite usar '
           'el material.'),
]
pasos_5_6 = [
    ('5.', 'Se lavarán las manos concienzudamente con agua y jabón antes de '
           'extraer el material del esterilizador o el recipiente en el que '
           'ha hervido. Se recomienda utilizar unas pinzas de cocina '
           'esterilizadas para manipular el material de preparación y '
           'administración esterilizado.'),
    ('6.', 'Para evitar la contaminación, lo más indicado es extraer el '
           'material justo antes de utilizarlo. Si el material se extrae y '
           'no se utiliza de inmediato, habrá que cubrirlo y guardarlo en un '
           'lugar limpio. Los biberones pueden ensamblarse por completo para '
           'impedir la contaminación del interior de la botella y el interior '
           'y el exterior de la tetina.'),
]


def paso_par(num, texto):
    return Paragraph(f'<b>{num}</b>  {texto}', step_style)


def subpaso_par(letra, texto):
    return Paragraph(f'{letra}  {texto}', substep_style)


def seccion(titulo, primer_paso):
    """Encabezado de sección junto a su primer paso (evita títulos huérfanos)."""
    return KeepTogether([
        Paragraph(f'<b>{titulo}</b>', h2_style),
        primer_paso,
    ])


credits = Paragraph(
    'El texto corresponde al apartado 3.1.1, «Limpieza y esterilización del '
    'material de preparación y administración». Imágenes aportadas por el '
    'autor.',
    note_style,
)


def footer(canvas, doc_):
    canvas.saveState()
    canvas.setFont('Times New Roman', 8)
    canvas.setFillColor(TEXT_MUTED)
    canvas.drawCentredString(
        A4[0] / 2.0, 1.1 * cm,
        'Limpieza y esterilización del material de preparación y '
        'administración · Apartado 3.1.1')
    canvas.restoreState()


story = [
    Paragraph('<b>Limpieza y esterilización del material de preparación '
              'y administración</b>', title_style),
    Paragraph('Apartado 3.1.1 · Procedimiento en 6 pasos', subtitle_style),
    HRFlowable(width='100%', thickness=1.2, color=ACCENT,
               spaceBefore=0, spaceAfter=14),
    intro_callout,
    Spacer(1, 4),
    seccion('Lavado de manos y limpieza (pasos 1 a 3)', paso_par(*pasos_1_3[0])),
    paso_par(*pasos_1_3[1]),
    paso_par(*pasos_1_3[2]),
    Spacer(1, 8),
    KeepTogether(fila_imagenes),
    seccion('Esterilización (paso 4)', paso_par(*paso_4)),
    subpaso_par(*subpasos[0]),
    subpaso_par(*subpasos[1]),
    subpaso_par(*subpasos[2]),
    seccion('Manipulación del material esterilizado (pasos 5 y 6)',
            paso_par(*pasos_5_6[0])),
    paso_par(*pasos_5_6[1]),
    Spacer(1, 8),
    KeepTogether(fila_imagenes_2),
    credits,
]

os.makedirs(OUT_DIR, exist_ok=True)
doc.build(story, onFirstPage=footer, onLaterPages=footer)
print('PDF generado:', OUT)
