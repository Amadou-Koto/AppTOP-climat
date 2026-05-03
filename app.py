"""Application Flask — Étude climatique multi-variables et comparative."""
import os
import pickle
import base64

from flask import (Flask, render_template, request, send_file,
                   session, redirect, url_for, flash)

from services.geocodage import geocoder
from services.meteo import (recuperer_donnees, recuperer_donnees_multi,
                            calculer_pivot, calculer_pivot_comparaison,
                            formater_tableau, formater_tableau_comparaison,
                            calculer_extremes, VARIABLES)
from services.exports import (export_word, export_excel, export_pdf,
                              export_word_comparaison,
                              export_excel_comparaison,
                              export_pdf_comparaison)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key-change-me")


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", variables=VARIABLES)


@app.route("/comparer", methods=["GET"])
def comparer():
    return render_template("comparer.html", variables=VARIABLES)


# ============================================================
# Mode mono-lieu
# ============================================================
@app.route("/analyser", methods=["POST"])
def analyser():
    mode = request.form.get("mode")
    nb_annees = int(request.form.get("nb_annees", 5))
    variables_actives = request.form.getlist("variables")

    if not variables_actives:
        flash("Veuillez sélectionner au moins une variable.", "error")
        return redirect(url_for("index"))

    try:
        if mode == "nom":
            nom = request.form.get("nom_lieu", "").strip()
            if not nom:
                flash("Veuillez saisir un nom de lieu.", "error")
                return redirect(url_for("index"))
            coords = geocoder(nom)
            lieu_affiche = nom
        else:
            lat = float(request.form.get("latitude"))
            lon = float(request.form.get("longitude"))
            coords = {"latitude": lat, "longitude": lon,
                      "nom_complet": f"{lat:.4f}° / {lon:.4f}°"}
            lieu_affiche = coords["nom_complet"]

        df = recuperer_donnees(coords["latitude"], coords["longitude"],
                                nb_annees, variables_actives)
        pivots = {v: calculer_pivot(df, v) for v in variables_actives}
        extremes = calculer_extremes(df, variables_actives)

        tableaux_html = []
        for var_key, pivot in pivots.items():
            tableaux_html.append({
                "cle": var_key,
                "config": VARIABLES[var_key],
                "entetes": ["Mois"] + [str(c) for c in pivot.columns],
                "lignes": formater_tableau(pivot),
            })

        session["mode_export"] = "mono"
        session["lieu"] = lieu_affiche
        session["coords"] = coords
        session["nb_annees"] = nb_annees
        session["pivots_pickle"] = base64.b64encode(pickle.dumps(pivots)).decode()
        session["extremes_pickle"] = base64.b64encode(pickle.dumps(extremes)).decode()

        return render_template("resultats.html",
                                lieu=lieu_affiche, coords=coords,
                                nb_annees=nb_annees,
                                tableaux=tableaux_html,
                                extremes=extremes)

    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("index"))
    except Exception as e:
        flash(f"Erreur : {e}", "error")
        return redirect(url_for("index"))


# ============================================================
# Mode comparaison
# ============================================================
@app.route("/analyser-comparaison", methods=["POST"])
def analyser_comparaison():
    nb_annees = int(request.form.get("nb_annees", 5))
    variables_actives = request.form.getlist("variables")
    noms_lieux = request.form.getlist("nom_lieu")
    noms_lieux = [n.strip() for n in noms_lieux if n.strip()]

    if not variables_actives:
        flash("Veuillez sélectionner au moins une variable.", "error")
        return redirect(url_for("comparer"))
    if len(noms_lieux) < 2:
        flash("Veuillez saisir au moins 2 lieux à comparer.", "error")
        return redirect(url_for("comparer"))

    try:
        lieux = []
        for nom in noms_lieux:
            coords = geocoder(nom)
            lieux.append({
                "nom": nom,
                "latitude": coords["latitude"],
                "longitude": coords["longitude"],
            })

        donnees_multi = recuperer_donnees_multi(lieux, nb_annees, variables_actives)
        pivots_compare = {
            v: calculer_pivot_comparaison(donnees_multi, v)
            for v in variables_actives
        }

        extremes_par_lieu = {}
        for lieu in lieux:
            df_lieu = donnees_multi[lieu["nom"]]
            extremes_par_lieu[lieu["nom"]] = calculer_extremes(df_lieu, variables_actives)

        resumes_lieux = []
        for lieu in lieux:
            df_lieu = donnees_multi[lieu["nom"]]
            resume = {
                "nom": lieu["nom"],
                "latitude": lieu["latitude"],
                "longitude": lieu["longitude"],
                "indicateurs": [],
            }
            for var_key in variables_actives:
                config = VARIABLES[var_key]
                if config["agg"] == "sum":
                    cumul_annuel = df_lieu.groupby("annee")[config["champ_api"]].sum().mean()
                    valeur = round(cumul_annuel, 1)
                    libelle = f"Cumul annuel {config['unite']}"
                else:
                    valeur = round(df_lieu[config["champ_api"]].mean(), 1)
                    libelle = f"Moyenne {config['unite']}"
                resume["indicateurs"].append({
                    "icone": config["icone"],
                    "label": config["label"].split(" (")[0],
                    "valeur": valeur,
                    "libelle": libelle,
                })
            resumes_lieux.append(resume)

        tableaux_html = []
        for var_key, pivot in pivots_compare.items():
            tableaux_html.append({
                "cle": var_key,
                "config": VARIABLES[var_key],
                "entetes": ["Mois"] + list(pivot.columns),
                "lignes": formater_tableau_comparaison(pivot),
            })

        session["mode_export"] = "compare"
        session["lieux_noms"] = noms_lieux
        session["nb_annees"] = nb_annees
        session["pivots_pickle"] = base64.b64encode(
            pickle.dumps(pivots_compare)).decode()
        session["extremes_pickle"] = base64.b64encode(
            pickle.dumps(extremes_par_lieu)).decode()

        return render_template("resultats_comparaison.html",
                                lieux_noms=noms_lieux,
                                lieux=lieux,
                                resumes_lieux=resumes_lieux,
                                extremes_par_lieu=extremes_par_lieu,
                                nb_annees=nb_annees,
                                tableaux=tableaux_html)

    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("comparer"))
    except Exception as e:
        flash(f"Erreur : {e}", "error")
        return redirect(url_for("comparer"))


# ============================================================
# Exports
# ============================================================
@app.route("/export/<format>")
def export(format):
    if "pivots_pickle" not in session:
        flash("Aucune donnée à exporter.", "error")
        return redirect(url_for("index"))

    pivots = pickle.loads(base64.b64decode(session["pivots_pickle"]))
    nb_annees = session["nb_annees"]
    mode_export = session.get("mode_export", "mono")
    extremes = pickle.loads(base64.b64decode(session["extremes_pickle"])) \
        if "extremes_pickle" in session else None

    mimes = {
        "word": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "excel": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "pdf": "application/pdf",
    }
    extensions = {"word": "docx", "excel": "xlsx", "pdf": "pdf"}

    if mode_export == "mono":
        lieu = session["lieu"]
        coords = session["coords"]
        if format == "word":
            buf = export_word(lieu, coords, nb_annees, pivots, extremes)
        elif format == "excel":
            buf = export_excel(lieu, coords, nb_annees, pivots, extremes)
        elif format == "pdf":
            buf = export_pdf(lieu, coords, nb_annees, pivots, extremes)
        else:
            flash("Format inconnu.", "error")
            return redirect(url_for("index"))
        nom_fichier = f"climat_{lieu}.{extensions[format]}"
    else:
        lieux_noms = session["lieux_noms"]
        if format == "word":
            buf = export_word_comparaison(lieux_noms, nb_annees, pivots, extremes)
        elif format == "excel":
            buf = export_excel_comparaison(lieux_noms, nb_annees, pivots, extremes)
        elif format == "pdf":
            buf = export_pdf_comparaison(lieux_noms, nb_annees, pivots, extremes)
        else:
            flash("Format inconnu.", "error")
            return redirect(url_for("index"))
        nom_fichier = f"comparaison_climat.{extensions[format]}"

    return send_file(buf, as_attachment=True,
                     download_name=nom_fichier, mimetype=mimes[format])


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", debug=False, port=port)
