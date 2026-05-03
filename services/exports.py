"""Génération des fichiers Word, Excel et PDF."""
import io
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, Image, PageBreak)

from services.meteo import MOIS_NOMS, VARIABLES


MOIS_COURTS = ["Jan", "Fév", "Mar", "Avr", "Mai", "Juin",
               "Juil", "Août", "Sep", "Oct", "Nov", "Déc"]


# ============================================================
# Graphiques
# ============================================================
def generer_graphique_ombrothermique(pivots: dict) -> io.BytesIO:
    """Diagramme classique température + pluie (si les deux sont présents)."""
    if "temperature" not in pivots or "pluie" not in pivots:
        return None

    fig, ax1 = plt.subplots(figsize=(10, 5))
    ax1.bar(MOIS_COURTS, pivots["pluie"]["Moyenne"].values,
            color="#1f77b4", alpha=0.6, label="Précipitations")
    ax1.set_xlabel("Mois")
    ax1.set_ylabel("Précipitations (mm)", color="#1f77b4")
    ax1.tick_params(axis="y", labelcolor="#1f77b4")

    ax2 = ax1.twinx()
    ax2.plot(MOIS_COURTS, pivots["temperature"]["Moyenne"].values,
             color="#d62728", marker="o", linewidth=2, label="Température")
    ax2.set_ylabel("Température (°C)", color="#d62728")
    ax2.tick_params(axis="y", labelcolor="#d62728")

    plt.title("Diagramme ombrothermique — Moyennes")
    fig.tight_layout()

    buffer = io.BytesIO()
    plt.savefig(buffer, format="png", dpi=150, bbox_inches="tight")
    plt.close()
    buffer.seek(0)
    return buffer


def generer_graphique_simple(pivot: pd.DataFrame, variable: str) -> io.BytesIO:
    """Graphique mensuel simple pour une variable secondaire."""
    config = VARIABLES[variable]
    fig, ax = plt.subplots(figsize=(10, 4))

    couleurs = {
        "humidite": "#17becf",
        "vent": "#7f7f7f",
        "ensoleillement": "#ff7f0e",
    }
    couleur = couleurs.get(variable, "#2ca02c")

    ax.bar(MOIS_COURTS, pivot["Moyenne"].values, color=couleur, alpha=0.75)
    ax.set_xlabel("Mois")
    ax.set_ylabel(config["label"])
    ax.set_title(f"{config['label']} — Moyenne mensuelle")
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    fig.tight_layout()

    buffer = io.BytesIO()
    plt.savefig(buffer, format="png", dpi=150, bbox_inches="tight")
    plt.close()
    buffer.seek(0)
    return buffer


def generer_graphique_comparaison(pivot: pd.DataFrame, variable: str) -> io.BytesIO:
    """Courbes superposées : une courbe par lieu."""
    config = VARIABLES[variable]
    fig, ax = plt.subplots(figsize=(11, 5))

    palette = ["#1f77b4", "#d62728", "#2ca02c", "#ff7f0e", "#9467bd"]

    for i, lieu in enumerate(pivot.columns):
        ax.plot(MOIS_COURTS, pivot[lieu].values,
                color=palette[i % len(palette)],
                marker="o", linewidth=2, label=lieu)

    ax.set_xlabel("Mois")
    ax.set_ylabel(config["label"])
    ax.set_title(f"Comparaison — {config['label']}")
    ax.legend(loc="best", framealpha=0.9)
    ax.grid(linestyle="--", alpha=0.4)
    fig.tight_layout()

    buffer = io.BytesIO()
    plt.savefig(buffer, format="png", dpi=150, bbox_inches="tight")
    plt.close()
    buffer.seek(0)
    return buffer


# ============================================================
# Helpers
# ============================================================
def _pivot_vers_lignes(pivot: pd.DataFrame):
    entetes = ["Mois"] + [str(c) for c in pivot.columns]
    lignes = []
    for mois_num in pivot.index:
        ligne = [MOIS_NOMS[mois_num - 1]]
        for col in pivot.columns:
            ligne.append(f"{pivot.loc[mois_num, col]:.1f}")
        lignes.append(ligne)
    return entetes, lignes


def _ajouter_tableau_word(doc, pivot):
    entetes, lignes = _pivot_vers_lignes(pivot)
    table = doc.add_table(rows=1, cols=len(entetes))
    table.style = "Light Grid Accent 1"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, txt in enumerate(entetes):
        cell = table.rows[0].cells[i]
        cell.text = txt
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True

    for ligne in lignes:
        cells = table.add_row().cells
        for i, val in enumerate(ligne):
            cells[i].text = val
            if i > 0:
                for para in cells[i].paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _ajouter_extremes_word(doc, extremes, titre_section="Statistiques d'extrêmes"):
    if not extremes:
        return
    doc.add_heading(titre_section, level=2)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Light Grid Accent 1"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    entetes = ["Indicateur", "Valeur", "Date / Période"]
    for i, txt in enumerate(entetes):
        cell = table.rows[0].cells[i]
        cell.text = txt
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for cle, e in extremes.items():
        row = table.add_row().cells
        row[0].text = f"{e['icone']} {e['label']}"
        row[1].text = e["valeur"]
        row[2].text = e["date"]
        for para in row[1].paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _ajouter_extremes_pdf(elements, extremes, style_h2,
                            titre_section="Statistiques d'extrêmes"):
    if not extremes:
        return
    elements.append(Paragraph(titre_section, style_h2))
    data = [["Indicateur", "Valeur", "Date / Période"]]
    for cle, e in extremes.items():
        data.append([f"{e['icone']} {e['label']}", e["valeur"], e["date"]])
    table = Table(data, repeatRows=1, colWidths=[8*cm, 5*cm, 9*cm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bol
