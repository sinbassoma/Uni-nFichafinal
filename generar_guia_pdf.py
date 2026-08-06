# -*- coding: utf-8 -*-
"""Genera la guía PDF de uso de la Ficha Social UNCP."""
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable,
)
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
import os

OUT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "Guia_Ficha_Social_UNCP.pdf",
)

NAVY = HexColor("#1a365d")
TEAL = HexColor("#0d9488")
LIGHT = HexColor("#f1f5f9")
SOFT = HexColor("#64748b")
DARK = HexColor("#0f172a")
BORDER = HexColor("#e2e8f0")

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(
    name="H1", fontName="Helvetica-Bold", fontSize=14,
    textColor=NAVY, spaceBefore=14, spaceAfter=8, leading=18,
))
styles.add(ParagraphStyle(
    name="H2", fontName="Helvetica-Bold", fontSize=11,
    textColor=TEAL, spaceBefore=10, spaceAfter=5, leading=14,
))
styles.add(ParagraphStyle(
    name="Body", fontName="Helvetica", fontSize=9.5,
    textColor=DARK, alignment=TA_JUSTIFY, spaceAfter=5, leading=13,
))
styles.add(ParagraphStyle(
    name="Item", fontName="Helvetica", fontSize=9.5,
    textColor=DARK, leftIndent=12, spaceAfter=3, leading=12.5,
))
styles.add(ParagraphStyle(
    name="Link", fontName="Helvetica", fontSize=8.2,
    textColor=HexColor("#0369a1"), spaceAfter=3, leading=11,
))
styles.add(ParagraphStyle(
    name="Small", fontName="Helvetica", fontSize=8,
    textColor=SOFT, spaceAfter=3, leading=10,
))
styles.add(ParagraphStyle(
    name="TableCell", fontName="Helvetica", fontSize=8.5,
    textColor=DARK, leading=11,
))
styles.add(ParagraphStyle(
    name="TableHead", fontName="Helvetica-Bold", fontSize=8.5,
    textColor=white, leading=11,
))


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.rect(0, A4[1] - 12 * mm, A4[0], 12 * mm, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont("Helvetica-Bold", 8)
    canvas.drawString(18 * mm, A4[1] - 7.5 * mm, "UNCP · Ficha Social — Guía de uso")
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(A4[0] - 18 * mm, A4[1] - 7.5 * mm, "2026")
    canvas.setStrokeColor(BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(18 * mm, 14 * mm, A4[0] - 18 * mm, 14 * mm)
    canvas.setFillColor(SOFT)
    canvas.setFont("Helvetica", 8)
    canvas.drawString(18 * mm, 8 * mm, "Unión Nacional de Ciegos del Perú")
    canvas.drawRightString(A4[0] - 18 * mm, 8 * mm, f"Pág. {doc.page}")
    canvas.restoreState()


def cover_page(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.rect(0, 0, A4[0], A4[1], fill=1, stroke=0)
    canvas.setFillColor(TEAL)
    canvas.rect(0, A4[1] * 0.38, A4[0], 3 * mm, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont("Helvetica-Bold", 11)
    canvas.drawCentredString(A4[0] / 2, A4[1] * 0.72, "UNIÓN NACIONAL DE CIEGOS DEL PERÚ")
    canvas.setFont("Helvetica-Bold", 24)
    canvas.drawCentredString(A4[0] / 2, A4[1] * 0.62, "Ficha Social UNCP")
    canvas.setFont("Helvetica", 13)
    canvas.drawCentredString(A4[0] / 2, A4[1] * 0.56, "Guía de uso y enlaces principales")
    canvas.setFont("Helvetica", 10)
    canvas.setFillColor(HexColor("#94a3b8"))
    canvas.drawCentredString(
        A4[0] / 2, A4[1] * 0.32,
        "Entrevista digital · PDF · Excel · Google Sheets · Drive",
    )
    canvas.setFont("Helvetica", 9)
    canvas.drawCentredString(A4[0] / 2, A4[1] * 0.18, "Documento para personal y administradores")
    canvas.drawCentredString(A4[0] / 2, A4[1] * 0.15, "Agosto 2026")
    canvas.restoreState()


def main():
    doc = SimpleDocTemplate(
        OUT,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
        title="Guía Ficha Social UNCP",
        author="UNCP",
    )
    story = []
    story.append(Spacer(1, 200 * mm))
    story.append(PageBreak())

    # 1
    story.append(Paragraph("1. ¿Qué es este sistema?", styles["H1"]))
    story.append(Paragraph(
        "La <b>Ficha Social UNCP</b> es una entrevista digital que reemplaza el llenado manual. "
        "El profesional responde las preguntas en pantalla (con o sin voz), y al final se genera "
        "la ficha oficial en <b>PDF</b> (mismo formato que se imprime), en <b>Excel</b> y se guarda "
        "en la nube: una <b>fila en Google Sheets</b> y los archivos en una <b>carpeta de Drive</b>.",
        styles["Body"],
    ))

    # 2
    story.append(Paragraph("2. Enlaces principales", styles["H1"]))
    story.append(Paragraph(
        "Guarde o marque estos enlaces. Son los destinos actuales del sistema:",
        styles["Body"],
    ))

    link_data = [
        [
            Paragraph("<b>Recurso</b>", styles["TableHead"]),
            Paragraph("<b>Enlace / uso</b>", styles["TableHead"]),
        ],
        [
            Paragraph("Hoja Google<br/>(filas de fichas)", styles["TableCell"]),
            Paragraph(
                "https://docs.google.com/spreadsheets/d/1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek/edit<br/>"
                '<font color="#64748b">Pestaña «Fichas»: datos + enlaces a PDF/Excel</font>',
                styles["Link"],
            ),
        ],
        [
            Paragraph("Carpeta Drive<br/>(PDF y Excel)", styles["TableCell"]),
            Paragraph(
                "https://drive.google.com/drive/folders/1dqy2HuWf6IYVGasuyrtff2id7Jf3aRXv<br/>"
                '<font color="#64748b">Aquí se guardan los archivos de cada entrevista</font>',
                styles["Link"],
            ),
        ],
        [
            Paragraph("Aplicación web<br/>(Apps Script)", styles["TableCell"]),
            Paragraph(
                "https://script.google.com/macros/s/AKfycbw2Gg0qOVeBjiRl0N0uNlgxsdUkbiPV8D1wDGepXv7H3TbXtfECszqNGI95IN0y17Az/exec<br/>"
                '<font color="#64748b">Puente entrevista ↔ Google. Debe mostrar JSON al abrirla en el navegador.</font>',
                styles["Link"],
            ),
        ],
        [
            Paragraph("Proyecto local<br/>(PC)", styles["TableCell"]),
            Paragraph(
                "Carpeta: uciego - sin audio final - copia<br/>"
                '<font color="#64748b">Abrir siempre con «Abrir entrevista.bat»</font>',
                styles["TableCell"],
            ),
        ],
    ]
    t = Table(link_data, colWidths=[42 * mm, 128 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("BACKGROUND", (0, 1), (-1, 1), LIGHT),
        ("BACKGROUND", (0, 3), (-1, 3), LIGHT),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(t)
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph(
        "<b>Nota:</b> Si la URL /exec da 404 o login, hay que volver a implementar el script "
        "como «Cualquier persona» y copiar la URL nueva. Detalle en CONFIGURAR-GOOGLE.md.",
        styles["Small"],
    ))

    # 3
    story.append(Paragraph("3. Cómo usar la entrevista (día a día)", styles["H1"]))
    story.append(Paragraph("3.1 Abrir el sistema", styles["H2"]))
    for b in [
        "Vaya a la carpeta del proyecto en el PC.",
        "Haga doble clic en <b>Abrir entrevista.bat</b> (no abra el HTML con doble clic).",
        "Se abrirá el navegador con la Ficha Social.",
    ]:
        story.append(Paragraph("• " + b, styles["Item"]))

    story.append(Paragraph("3.2 Iniciar la ficha", styles["H2"]))
    for b in [
        "<b>Comenzar sin voz</b> (recomendado): solo pantalla.",
        "<b>Comenzar con voz</b> (opcional): el sistema lee cada pregunta.",
        "Si hay borrador, puede continuar donde quedó.",
    ]:
        story.append(Paragraph("• " + b, styles["Item"]))

    story.append(Paragraph("3.3 Durante la entrevista", styles["H2"]))
    for b in [
        "Responda cada pregunta (teclado, opciones o micrófono si está activo).",
        "Use <b>Anterior</b> / <b>Siguiente</b>.",
        "Las respuestas se guardan en borrador en el navegador.",
        "Al final verá el <b>resumen</b>; puede editar cualquier respuesta.",
    ]:
        story.append(Paragraph("• " + b, styles["Item"]))

    story.append(Paragraph("3.4 Botones al finalizar", styles["H2"]))
    btn_data = [
        [
            Paragraph("<b>Botón</b>", styles["TableHead"]),
            Paragraph("<b>Qué hace</b>", styles["TableHead"]),
        ],
        [
            Paragraph("☁ Guardar ficha en la nube", styles["TableCell"]),
            Paragraph(
                "Guarda la fila en la hoja y sube PDF + Excel (+ Word) a Drive. Paso principal.",
                styles["TableCell"],
            ),
        ],
        [
            Paragraph("📄 Descargar PDF", styles["TableCell"]),
            Paragraph(
                "Descarga en el PC la ficha con el <b>mismo formato que se imprime</b>.",
                styles["TableCell"],
            ),
        ],
        [
            Paragraph("🖨 Imprimir ficha", styles["TableCell"]),
            Paragraph(
                "Impresión del navegador (puede elegir «Guardar como PDF» del sistema).",
                styles["TableCell"],
            ),
        ],
        [
            Paragraph("📊 Descargar Excel", styles["TableCell"]),
            Paragraph(
                "Descarga Excel con la misma información de la ficha (plantilla oficial).",
                styles["TableCell"],
            ),
        ],
        [
            Paragraph("🗒 Nueva ficha", styles["TableCell"]),
            Paragraph("Limpia y empieza otra entrevista.", styles["TableCell"]),
        ],
    ]
    t2 = Table(btn_data, colWidths=[48 * mm, 122 * mm])
    t2.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), TEAL),
        ("BACKGROUND", (0, 2), (-1, 2), LIGHT),
        ("BACKGROUND", (0, 4), (-1, 4), LIGHT),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(t2)

    story.append(PageBreak())

    # 4
    story.append(Paragraph("4. Panel administrador", styles["H1"]))
    story.append(Paragraph(
        "No se muestra al socio. En la pantalla de resumen: <b>5 clics rápidos</b> en el pie "
        "«UNCP · Ficha social».",
        styles["Body"],
    ))
    for b in [
        "<b>Diagnosticar</b>: comprueba si el script de Google responde.",
        "<b>URL / config</b>: pegar la URL que termina en <b>/exec</b>.",
        "<b>↗ Hoja</b> y <b>↗ Carpeta</b>: atajos a Sheets y Drive.",
    ]:
        story.append(Paragraph("• " + b, styles["Item"]))

    # 5
    story.append(Paragraph("5. Configurar Google (si no guarda)", styles["H1"]))
    story.append(Paragraph(
        "Si aparece «Error de red» o «Failed to fetch», casi siempre es la "
        "<b>implementación del script</b>, no el internet del local.",
        styles["Body"],
    ))
    story.append(Paragraph("5.1 Script en la hoja", styles["H2"]))
    for i, b in enumerate([
        "Hoja → <b>Extensiones → Apps Script</b>.",
        "Pegue todo <b>apps-script/Codigo.gs</b> y guarde (Ctrl+S).",
        "Verifique SPREADSHEET_ID y PDF_FOLDER_ID de la hoja y carpeta actuales.",
    ], 1):
        story.append(Paragraph(f"<b>{i}.</b> {b}", styles["Item"]))

    story.append(Paragraph("5.2 Publicar aplicación web", styles["H2"]))
    for i, b in enumerate([
        "<b>Implementar → Nueva implementación</b> (o Administrar → lápiz → Nueva versión).",
        "Tipo: <b>Aplicación web</b>. Ejecutar como: <b>Yo</b>.",
        "Quién tiene acceso: <b>Cualquier persona</b> (obligatorio).",
        "Implementar → copiar URL <b>/exec</b>.",
        "Probar en <b>incógnito</b>: debe verse JSON con \"ok\": true (no login).",
        "En la entrevista: Admin → pegar /exec → Guardar → Diagnosticar.",
    ], 1):
        story.append(Paragraph(f"<b>{i}.</b> {b}", styles["Item"]))

    # 6
    story.append(Paragraph("6. Problemas frecuentes", styles["H1"]))
    prob = [
        [
            Paragraph("<b>Síntoma</b>", styles["TableHead"]),
            Paragraph("<b>Qué revisar</b>", styles["TableHead"]),
        ],
        [
            Paragraph("Error de red / Failed to fetch", styles["TableCell"]),
            Paragraph(
                "Apps Script: acceso «Cualquier persona»; URL /exec válida; en incógnito JSON.",
                styles["TableCell"],
            ),
        ],
        [
            Paragraph("404 en la URL /exec", styles["TableCell"]),
            Paragraph(
                "URL mal copiada o implementación borrada. Nueva implementación y copiar de nuevo.",
                styles["TableCell"],
            ),
        ],
        [
            Paragraph("401 o pantalla de login", styles["TableCell"]),
            Paragraph(
                "La app no es pública. Cambiar a Cualquier persona + Nueva versión.",
                styles["TableCell"],
            ),
        ],
        [
            Paragraph("Fila sí, PDF no", styles["TableCell"]),
            Paragraph(
                "Comparta la carpeta Drive como Editor con la cuenta del script.",
                styles["TableCell"],
            ),
        ],
        [
            Paragraph("Pantalla en blanco / sin micrófono", styles["TableCell"]),
            Paragraph(
                "Usar siempre <b>Abrir entrevista.bat</b>, no abrir el HTML a doble clic.",
                styles["TableCell"],
            ),
        ],
    ]
    t3 = Table(prob, colWidths=[48 * mm, 122 * mm])
    t3.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HexColor("#b45309")),
        ("BACKGROUND", (0, 2), (-1, 2), HexColor("#fffbeb")),
        ("BACKGROUND", (0, 4), (-1, 4), HexColor("#fffbeb")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(t3)

    story.append(Spacer(1, 8 * mm))
    story.append(HRFlowable(width="100%", thickness=1, color=BORDER, spaceAfter=6))
    story.append(Paragraph(
        "Archivos de ayuda: CONFIGURAR-GOOGLE.md · apps-script/Codigo.gs · "
        "Abrir entrevista.bat · INSTALAR-SCRIPT-HOJA.bat",
        styles["Small"],
    ))
    story.append(Paragraph(
        "© UNCP — Ficha Social digital. Guía de uso del sistema local y conexión a Google.",
        styles["Small"],
    ))

    doc.build(story, onFirstPage=cover_page, onLaterPages=header_footer)
    print("OK:", OUT)
    print("Size:", os.path.getsize(OUT))


if __name__ == "__main__":
    main()
