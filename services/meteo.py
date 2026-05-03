"""Récupération et agrégation des données météo via Open-Meteo (ERA5)."""
import locale
import requests
import pandas as pd
from datetime import date

# Tentative d'activation de la locale française pour les noms de mois
try:
    locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, "fr_FR")
    except locale.Error:
        pass


# Variables disponibles avec leur configuration
VARIABLES = {
    "temperature": {
        "label": "Températures moyennes (°C)",
        "champ_api": "temperature_2m_mean",
        "agg": "mean",
        "unite": "°C",
        "icone": "🌡️",
    },
    "pluie": {
        "label": "Précipitations (mm)",
        "champ_api": "precipitation_sum",
        "agg": "sum",
        "unite": "mm",
        "icone": "🌧️",
    },
    "humidite": {
        "label": "Humidité relative moyenne (%)",
        "champ_api": "relative_humidity_2m_mean",
        "agg": "mean",
        "unite": "%",
        "icone": "💧",
    },
    "vent": {
        "label": "Vitesse moyenne du vent (km/h)",
        "champ_api": "wind_speed_10m_max",
        "agg": "mean",
        "unite": "km/h",
        "icone": "💨",
    },
    "ensoleillement": {
        "label": "Ensoleillement (heures/jour)",
        "champ_api": "sunshine_duration",
        "agg": "mean",
        "unite": "h/j",
        "icone": "☀️",
    },
}

MOIS_NOMS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
             "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]


def recuperer_donnees(latitude: float, longitude: float, nb_annees: int,
                       variables_actives: list) -> pd.DataFrame:
    """
    Récupère les données journalières des n dernières années complètes.
    On ajoute systématiquement temperature_2m_max et _min pour les extrêmes.
    """
    annee_fin = date.today().year - 1
    annee_debut = annee_fin - nb_annees + 1

    champs_api = set(VARIABLES[v]["champ_api"] for v in variables_actives)
    champs_api.update(["temperature_2m_max", "temperature_2m_min"])

    url = "https://archive-api.open-meteo.com/v1/archive"
    params = {
        "latitude": latitude,
        "longitude": longitude,
        "start_date": f"{annee_debut}-01-01",
        "end_date": f"{annee_fin}-12-31",
        "daily": ",".join(champs_api),
        "timezone": "auto",
    }

    r = requests.get(url, params=params, timeout=30)
    r.raise_for_status()
    data = r.json()["daily"]

    df = pd.DataFrame(data)
    df["time"] = pd.to_datetime(df["time"])
    df["annee"] = df["time"].dt.year
    df["mois"] = df["time"].dt.month

    if "sunshine_duration" in df.columns:
        df["sunshine_duration"] = df["sunshine_duration"] / 3600

    return df


def calculer_pivot(df: pd.DataFrame, variable: str) -> pd.DataFrame:
    """Pivot mensuel pour une variable : lignes = mois, colonnes = années + Moyenne."""
    config = VARIABLES[variable]
    pivot = df.pivot_table(
        index="mois", columns="annee",
        values=config["champ_api"], aggfunc=config["agg"]
    ).round(1)
    pivot["Moyenne"] = pivot.mean(axis=1).round(1)
    return pivot


def formater_tableau(pivot: pd.DataFrame) -> list:
    """Convertit le pivot en liste de dicts pour Jinja."""
    lignes = []
    for mois_num in pivot.index:
        ligne = {"mois": MOIS_NOMS[mois_num - 1]}
        for col in pivot.columns:
            ligne[str(col)] = pivot.loc[mois_num, col]
        lignes.append(ligne)
    return lignes


def recuperer_donnees_multi(lieux: list, nb_annees: int,
                             variables_actives: list) -> dict:
    """Récupère les données pour plusieurs lieux. Retourne {nom_lieu: DataFrame}."""
    resultats = {}
    for lieu in lieux:
        df = recuperer_donnees(lieu["latitude"], lieu["longitude"],
                                nb_annees, variables_actives)
        resultats[lieu["nom"]] = df
    return resultats


def calculer_pivot_comparaison(donnees_multi: dict, variable: str) -> pd.DataFrame:
    """Pivot pour comparaison : lignes = mois, colonnes = lieux."""
    config = VARIABLES[variable]
    resultat = pd.DataFrame(index=range(1, 13))

    for nom_lieu, df in donnees_multi.items():
        if config["agg"] == "sum":
            par_mois_annee = df.groupby(["annee", "mois"])[config["champ_api"]].sum()
            serie = par_mois_annee.groupby("mois").mean()
        else:
            serie = df.groupby("mois")[config["champ_api"]].mean()
        resultat[nom_lieu] = serie.round(1)

    return resultat


def formater_tableau_comparaison(pivot: pd.DataFrame) -> list:
    """Formate le pivot de comparaison pour l'affichage HTML."""
    lignes = []
    for mois_num in pivot.index:
        ligne = {"mois": MOIS_NOMS[mois_num - 1]}
        for col in pivot.columns:
            ligne[str(col)] = pivot.loc[mois_num, col]
        lignes.append(ligne)
    return lignes


def calculer_extremes(df: pd.DataFrame, variables_actives: list) -> dict:
    """Calcule les statistiques d'extrêmes climatiques sur la période."""
    extremes = {}

    if "temperature_2m_max" in df.columns:
        idx_max = df["temperature_2m_max"].idxmax()
        extremes["temp_max"] = {
            "icone": "🔥",
            "label": "Jour le plus chaud",
            "valeur": f"{df.loc[idx_max, 'temperature_2m_max']:.1f} °C",
            "date": df.loc[idx_max, "time"].strftime("%d %B %Y"),
        }

    if "temperature_2m_min" in df.columns:
        idx_min = df["temperature_2m_min"].idxmin()
        extremes["temp_min"] = {
            "icone": "❄️",
            "label": "Jour le plus froid",
            "valeur": f"{df.loc[idx_min, 'temperature_2m_min']:.1f} °C",
            "date": df.loc[idx_min, "time"].strftime("%d %B %Y"),
        }

    if "pluie" in variables_actives and "precipitation_sum" in df.columns:
        idx_pluie = df["precipitation_sum"].idxmax()
        extremes["pluie_max_jour"] = {
            "icone": "🌊",
            "label": "Pluie max en 24h",
            "valeur": f"{df.loc[idx_pluie, 'precipitation_sum']:.1f} mm",
            "date": df.loc[idx_pluie, "time"].strftime("%d %B %Y"),
        }

        cumul_par_mois = df.groupby([df["time"].dt.year, df["time"].dt.month])[
            "precipitation_sum"
        ].sum()
        idx_mois_max = cumul_par_mois.idxmax()
        annee, mois = idx_mois_max
        extremes["mois_plus_pluvieux"] = {
            "icone": "🌧️",
            "label": "Mois le plus pluvieux",
            "valeur": f"{cumul_par_mois.max():.1f} mm",
            "date": f"{MOIS_NOMS[mois - 1]} {annee}",
        }

        jours_secs = (df["precipitation_sum"] < 1.0).astype(int)
        groupes = (jours_secs != jours_secs.shift()).cumsum()
        longueurs = jours_secs.groupby(groupes).sum()
        max_serie = longueurs.max()
        idx_groupe_max = longueurs.idxmax()
        date_fin_serie = df.loc[groupes == idx_groupe_max, "time"].iloc[-1]
        extremes["plus_longue_secheresse"] = {
            "icone": "🏜️",
            "label": "Plus longue période sèche",
            "valeur": f"{int(max_serie)} jours",
            "date": f"se terminant le {date_fin_serie.strftime('%d %B %Y')}",
        }

    if "vent" in variables_actives and "wind_speed_10m_max" in df.columns:
        idx_vent = df["wind_speed_10m_max"].idxmax()
        extremes["vent_max"] = {
            "icone": "💨",
            "label": "Jour le plus venté",
            "valeur": f"{df.loc[idx_vent, 'wind_speed_10m_max']:.1f} km/h",
            "date": df.loc[idx_vent, "time"].strftime("%d %B %Y"),
        }

    if "humidite" in variables_actives and "relative_humidity_2m_mean" in df.columns:
        idx_h_max = df["relative_humidity_2m_mean"].idxmax()
        idx_h_min = df["relative_humidity_2m_mean"].idxmin()
        extremes["humidite_max"] = {
            "icone": "💧",
            "label": "Jour le plus humide",
            "valeur": f"{df.loc[idx_h_max, 'relative_humidity_2m_mean']:.0f} %",
            "date": df.loc[idx_h_max, "time"].strftime("%d %B %Y"),
        }
        extremes["humidite_min"] = {
            "icone": "🌵",
            "label": "Jour le plus sec",
            "valeur": f"{df.loc[idx_h_min, 'relative_humidity_2m_mean']:.0f} %",
            "date": df.loc[idx_h_min, "time"].strftime("%d %B %Y"),
        }

    if "ensoleillement" in variables_actives and "sunshine_duration" in df.columns:
        idx_soleil = df["sunshine_duration"].idxmax()
        extremes["ensoleillement_max"] = {
            "icone": "☀️",
            "label": "Journée la plus ensoleillée",
            "valeur": f"{df.loc[idx_soleil, 'sunshine_duration']:.1f} h",
            "date": df.loc[idx_soleil, "time"].strftime("%d %B %Y"),
        }

    return extremes
