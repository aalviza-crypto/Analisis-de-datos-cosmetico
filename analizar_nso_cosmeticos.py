#!/usr/bin/env python3
"""
Analiza las Notificaciones Sanitarias Obligatorias (NSO) de cosméticos del INVIMA
a partir del dataset abierto "INFORMACIÓN SOBRE RS Y NSO VIGENTES Y VENCIDAS"
(datos.gov.co, id ui32-p9f2).

Salida:
  - Ranking de formas cosméticas más notificadas (crema, champú, loción, gel...).
  - Ranking de usos / indicaciones cosméticas (capilar, facial, corporal, labial...).
  - Cruce forma x uso.
  - Top titulares.
  Todo en un Excel (resultados_nso_cosmeticos.xlsx) y un resumen por consola.

La forma y el uso NO son campos del dataset: se infieren del nombre comercial del
producto (campo `producto`) con palabras clave. Por eso el resultado es una
aproximación; el porcentaje "sin clasificar" se reporta para que se vea la cobertura.

Uso:
  python analizar_nso_cosmeticos.py                 # descarga desde la API Socrata
  python analizar_nso_cosmeticos.py --csv datos.csv # usa un CSV descargado a mano
  python analizar_nso_cosmeticos.py --solo-vigentes
Opcional: variable de entorno SOCRATA_APP_TOKEN para evitar límites de la API.
"""
import argparse
import os
import re
import sys
import unicodedata

import pandas as pd
import requests

API = "https://www.datos.gov.co/resource/ui32-p9f2.json"
PAGE = 50000

# Orden importa: la primera coincidencia gana (lo más específico primero).
FORMAS = [
    ("Aerosol / spray", r"\b(aerosol|spray|bruma|mist)\b"),
    ("Champú", r"\b(champu|shampoo|shampu)\b"),
    ("Acondicionador", r"\b(acondicionador|conditioner|enjuague capilar)\b"),
    ("Jabón", r"\b(jabon|soap|syndet)\b"),
    ("Gel", r"\b(gel|gelatina|jelly)\b"),
    ("Crema", r"\b(crema|cream|creme|cold cream)\b"),
    ("Emulsión / leche", r"\b(emulsion|leche|milk|lotion cream|body milk)\b"),
    ("Loción", r"\b(locion|lotion)\b"),
    ("Tónico / agua micelar", r"\b(tonico|tonic|micelar|agua de rosas)\b"),
    ("Sérum / ampolla", r"\b(serum|suero|ampolla|ampoule|concentrado)\b"),
    ("Aceite", r"\b(aceite|oil|oleo)\b"),
    ("Mascarilla", r"\b(mascarilla|mask|masque|tratamiento capilar)\b"),
    ("Exfoliante", r"\b(exfoliante|scrub|peeling)\b"),
    ("Bálsamo / manteca", r"\b(balsamo|balm|manteca|butter)\b"),
    ("Pomada / cera", r"\b(pomada|cera|wax|pasta)\b"),
    ("Labial / barra", r"\b(labial|lipstick|barra|stick|brillo labial|lip gloss|gloss)\b"),
    ("Polvo", r"\b(polvo|powder|talco|rubor|blush|sombra|compacto)\b"),
    ("Base / maquillaje líquido", r"\b(base|foundation|corrector|concealer|bb cream|cc cream)\b"),
    ("Esmalte", r"\b(esmalte|nail polish|laca|top coat|base coat)\b"),
    ("Perfume / colonia", r"\b(perfume|colonia|eau de|fragancia|parfum|cologne|splash)\b"),
    ("Desodorante / antitranspirante", r"\b(desodorante|antitranspirante|deodorant|roll on)\b"),
    ("Tinte / coloración", r"\b(tinte|coloracion|color crema|decolorante|oxigenta|oxidante|peroxido)\b"),
    ("Espuma / mousse", r"\b(espuma|mousse|foam)\b"),
    ("Pestañina / delineador", r"\b(pestanina|mascara de pestanas|rimel|delineador|eyeliner|lapiz)\b"),
    ("Toallitas", r"\b(toallita|toallitas|wipes|panitos)\b"),
    ("Crema dental / enjuague bucal", r"\b(crema dental|dental|dentifrico|enjuague bucal|colutorio|toothpaste)\b"),
]

USOS = [
    ("Protección solar", r"\b(solar|bloqueador|protector solar|sunscreen|spf|fps|bronceador|autobronceador)\b"),
    ("Capilar (lavado / acondicionamiento / tratamiento)", r"\b(champu|shampoo|acondicionador|capilar|cabello|pelo|hair|keratina|caspa|rizos|anticaida|caida)\b"),
    ("Coloración / decoloración capilar", r"\b(tinte|coloracion|decolorante|oxigenta|peroxido|matizante)\b"),
    ("Alisado / fijación / peinado", r"\b(alisado|alisante|fijador|gel fijador|laca|cera|pomada|styling|permanente|ondulado)\b"),
    ("Higiene oral", r"\b(dental|dentifrico|bucal|colutorio|toothpaste)\b"),
    ("Desodorante / antitranspirante", r"\b(desodorante|antitranspirante|deodorant)\b"),
    ("Uñas", r"\b(esmalte|unas|nail|cuticula|quitaesmalte|removedor)\b"),
    ("Maquillaje", r"\b(maquillaje|base|labial|rubor|sombra|pestanina|delineador|corrector|polvo compacto|lip|gloss|makeup)\b"),
    ("Perfumería / fragancia", r"\b(perfume|colonia|eau de|fragancia|parfum|cologne|splash)\b"),
    ("Afeitado / cuidado masculino", r"\b(afeitar|afeitado|shave|after shave|barba|beard|men|hombre)\b"),
    ("Despigmentante / aclarante", r"\b(despigmentante|aclarante|blanqueador|manchas|whitening|kojico|niacinamida)\b"),
    ("Antiedad / antiarrugas", r"\b(antiedad|antiarrugas|anti edad|anti age|antiage|retinol|colageno|hialuronico|reafirmante)\b"),
    ("Anti-acné / piel grasa", r"\b(acne|antiacne|matificante|piel grasa|salicilico|seborreguladora)\b"),
    ("Limpieza facial / desmaquillante", r"\b(limpiador|limpiadora|desmaquillante|micelar|cleanser|facial wash)\b"),
    ("Exfoliación", r"\b(exfoliante|scrub|peeling)\b"),
    ("Hidratación corporal / facial", r"\b(hidratante|humectante|moisturizing|nutritiva|corporal|body|manos|pies)\b"),
    ("Higiene corporal (baño)", r"\b(jabon|bano|ducha|shower|intimo|intima)\b"),
    ("Depilación", r"\b(depilatoria|depilatorio|depilacion|cera depilatoria)\b"),
    ("Cuidado infantil", r"\b(bebe|baby|ninos|kids|infantil)\b"),
    ("Repelente de insectos (cosmético)", r"\b(repelente)\b"),
]

_FORMAS_RE = [(n, re.compile(p)) for n, p in FORMAS]
_USOS_RE = [(n, re.compile(p)) for n, p in USOS]


def normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", texto.lower())


def clasificar(nombre, reglas):
    t = normalizar(nombre)
    for etiqueta, rx in reglas:
        if rx.search(t):
            return etiqueta
    return "Sin clasificar"


def descargar():
    headers = {}
    if os.environ.get("SOCRATA_APP_TOKEN"):
        headers["X-App-Token"] = os.environ["SOCRATA_APP_TOKEN"]
    where = "upper(grupo) like '%COSMET%'"
    filas, offset = [], 0
    while True:
        params = {"$where": where, "$limit": PAGE, "$offset": offset, "$order": ":id"}
        r = requests.get(API, params=params, headers=headers, timeout=120)
        r.raise_for_status()
        lote = r.json()
        filas.extend(lote)
        print(f"  descargados {len(filas)} registros...", file=sys.stderr)
        if len(lote) < PAGE:
            break
        offset += PAGE
    return pd.DataFrame(filas)


def cargar_csv(ruta):
    df = pd.read_csv(ruta, dtype=str, low_memory=False)
    df.columns = [normalizar(c).replace(" ", "") for c in df.columns]
    if "grupo" in df.columns:
        df = df[df["grupo"].map(normalizar).str.contains("cosmet", na=False)]
    return df


def analizar(df, solo_vigentes=False):
    df = df.copy()
    if solo_vigentes and "estadoregistro" in df.columns:
        df = df[df["estadoregistro"].map(normalizar).str.contains("vigente", na=False)]
    # El dataset repite cada NSO por cada rol (titular, fabricante, importador...):
    # contamos notificaciones únicas.
    clave = [c for c in ("rsynso", "producto") if c in df.columns]
    nso = df.drop_duplicates(subset=clave).copy()
    nso["forma"] = nso["producto"].map(lambda x: clasificar(x, _FORMAS_RE))
    nso["uso"] = nso["producto"].map(lambda x: clasificar(x, _USOS_RE))

    def ranking(col):
        s = nso[col].value_counts().rename_axis(col).reset_index(name="nso")
        s["%"] = (100 * s["nso"] / len(nso)).round(2)
        return s

    res = {
        "formas": ranking("forma"),
        "usos": ranking("uso"),
        "forma_x_uso": pd.crosstab(nso["forma"], nso["uso"]),
        "nso": nso,
    }
    if "titular" in nso.columns:
        res["titulares"] = nso["titular"].value_counts().head(50).rename_axis("titular").reset_index(name="nso")
    if "estadoregistro" in nso.columns:
        res["estado"] = nso["estadoregistro"].value_counts().rename_axis("estado").reset_index(name="nso")
    return res


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", help="CSV descargado de datos.gov.co (ui32-p9f2)")
    ap.add_argument("--solo-vigentes", action="store_true", help="Solo NSO en estado vigente")
    ap.add_argument("--salida", default="resultados_nso_cosmeticos.xlsx")
    args = ap.parse_args()

    df = cargar_csv(args.csv) if args.csv else descargar()
    if df.empty:
        sys.exit("No se obtuvieron registros de cosméticos.")
    res = analizar(df, args.solo_vigentes)

    total = len(res["nso"])
    print(f"\nNSO únicas de cosméticos analizadas: {total}\n")
    print("Formas cosméticas más notificadas:")
    print(res["formas"].head(20).to_string(index=False))
    print("\nUsos / indicaciones cosméticas más frecuentes:")
    print(res["usos"].head(20).to_string(index=False))

    with pd.ExcelWriter(args.salida) as xw:
        res["formas"].to_excel(xw, sheet_name="Formas", index=False)
        res["usos"].to_excel(xw, sheet_name="Usos", index=False)
        res["forma_x_uso"].to_excel(xw, sheet_name="Forma x Uso")
        for hoja in ("titulares", "estado"):
            if hoja in res:
                res[hoja].to_excel(xw, sheet_name=hoja.capitalize(), index=False)
        res["nso"].to_excel(xw, sheet_name="Detalle NSO", index=False)
    print(f"\nExcel generado: {args.salida}")


if __name__ == "__main__":
    main()
