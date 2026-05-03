"""Conversion d'un nom de lieu en coordonnées GPS via Nominatim (OpenStreetMap)."""
import requests


def geocoder(nom_lieu: str) -> dict:
    """
    Convertit un nom de lieu en coordonnées (latitude, longitude).
    Retourne un dict avec lat, lon et nom_complet, ou lève ValueError.
    """
    url = "https://nominatim.openstreetmap.org/search"
    params = {"q": nom_lieu, "format": "json", "limit": 1}
    headers = {"User-Agent": "ClimatApp/1.0"}  # exigé par Nominatim

    r = requests.get(url, params=params, headers=headers, timeout=10)
    r.raise_for_status()
    resultats = r.json()

    if not resultats:
        raise ValueError(f"Lieu introuvable : {nom_lieu}")

    return {
        "latitude": float(resultats[0]["lat"]),
        "longitude": float(resultats[0]["lon"]),
        "nom_complet": resultats[0]["display_name"],
    }
