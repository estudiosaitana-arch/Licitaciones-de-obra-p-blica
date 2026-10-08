#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vigilancia de licitaciones de obras en la Plataforma de Contratación del
Sector Público (PLACSP). Versión 2.

Qué hace en cada ejecución
  1. Lee el feed ATOM de la Plataforma (y, si se pide, los ficheros ZIP mensuales).
  2. Guarda en una base local (base_licitaciones.json) todas las licitaciones de
     obras del ámbito geográfico, sin duplicados, conservando la versión más reciente
     y anotando los cambios de estado, plazo e importe.
  3. Genera licitaciones_en_plazo.csv con las que siguen abiertas, ordenadas por
     fecha límite, marcando las nuevas desde la ejecución anterior.

Uso
  python licitaciones_obras.py                      # actualización (diaria)
  python licitaciones_obras.py --meses 2            # carga inicial con los ZIP mensuales
  python licitaciones_obras.py --min 500000 --max 10000000
  python licitaciones_obras.py --nuts ES521         # solo provincia de Alicante

Solo usa la biblioteca estándar de Python (3.8+).
"""

import argparse
import csv
import json
import os
import re
import sys
import time
import unicodedata
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime, timedelta
from urllib.parse import urljoin

FEED_URL = ("https://contrataciondelestado.es/sindicacion/sindicacion_643/"
            "licitacionesPerfilesContratanteCompleto3.atom")
ZIP_URL = ("https://contrataciondelsectorpublico.gob.es/sindicacion/sindicacion_643/"
           "licitacionesPerfilesContratanteCompleto3_{}.zip")
ATOM = "{http://www.w3.org/2005/Atom}"
BASE_POR_DEFECTO = "base_licitaciones.json"
SALIDA_POR_DEFECTO = "licitaciones_en_plazo.csv"

# Listas de códigos de la Plataforma (si aparece un código no recogido, se exporta tal cual)
PROCEDIMIENTOS = {
    "1": "Abierto", "2": "Restringido", "3": "Negociado sin publicidad",
    "4": "Negociado con publicidad", "5": "Diálogo competitivo", "6": "Contrato menor",
    "7": "Derivado de acuerdo marco", "8": "Concurso de proyectos", "9": "Abierto simplificado",
    "10": "Asociación para la innovación", "11": "Derivado de asociación para la innovación",
    "12": "Basado en sistema dinámico de adquisición", "13": "Licitación con negociación",
    "100": "Normas internas", "999": "Otros",
}
GARANTIAS = {"1": "provisional", "2": "definitiva", "3": "complementaria"}
UNIDADES = {"DAY": "días", "MON": "meses", "ANN": "años"}

COMUNIDADES = {
    "ES11": "Galicia", "ES12": "Principado de Asturias", "ES13": "Cantabria",
    "ES21": "País Vasco", "ES22": "Comunidad Foral de Navarra", "ES23": "La Rioja",
    "ES24": "Aragón", "ES30": "Comunidad de Madrid", "ES41": "Castilla y León",
    "ES42": "Castilla-La Mancha", "ES43": "Extremadura", "ES51": "Cataluña",
    "ES52": "Comunitat Valenciana", "ES53": "Illes Balears", "ES61": "Andalucía",
    "ES62": "Región de Murcia", "ES63": "Ceuta", "ES64": "Melilla", "ES70": "Canarias",
}
PROVINCIAS = {
    "ES111": "A Coruña", "ES112": "Lugo", "ES113": "Ourense", "ES114": "Pontevedra",
    "ES120": "Asturias", "ES130": "Cantabria", "ES211": "Araba/Álava", "ES212": "Gipuzkoa",
    "ES213": "Bizkaia", "ES220": "Navarra", "ES230": "La Rioja", "ES241": "Huesca",
    "ES242": "Teruel", "ES243": "Zaragoza", "ES300": "Madrid", "ES411": "Ávila",
    "ES412": "Burgos", "ES413": "León", "ES414": "Palencia", "ES415": "Salamanca",
    "ES416": "Segovia", "ES417": "Soria", "ES418": "Valladolid", "ES419": "Zamora",
    "ES421": "Albacete", "ES422": "Ciudad Real", "ES423": "Cuenca", "ES424": "Guadalajara",
    "ES425": "Toledo", "ES431": "Badajoz", "ES432": "Cáceres", "ES511": "Barcelona",
    "ES512": "Girona", "ES513": "Lleida", "ES514": "Tarragona", "ES521": "Alicante",
    "ES522": "Castellón", "ES523": "Valencia", "ES531": "Eivissa y Formentera",
    "ES532": "Mallorca", "ES533": "Menorca", "ES611": "Almería", "ES612": "Cádiz",
    "ES613": "Córdoba", "ES614": "Granada", "ES615": "Huelva", "ES616": "Jaén",
    "ES617": "Málaga", "ES618": "Sevilla", "ES620": "Murcia", "ES630": "Ceuta",
    "ES640": "Melilla", "ES703": "El Hierro", "ES704": "Fuerteventura",
    "ES705": "Gran Canaria", "ES706": "La Gomera", "ES707": "La Palma",
    "ES708": "Lanzarote", "ES709": "Tenerife",
}

CAMPOS_CSV = [
    "nueva", "expediente", "organo_contratacion", "objeto", "presupuesto_base_sin_iva",
    "valor_estimado", "fecha_limite_presentacion", "dias_restantes", "comunidad",
    "provincia", "lugar_ejecucion", "procedimiento", "plazo_ejecucion",
    "clasificacion_exigida", "criterios_adjudicacion", "pct_juicio_valor", "garantias",
    "revision_precios", "num_lotes", "cpv", "url_pcap", "url_ppt", "url_licitacion",
    "primera_deteccion", "cambios",
]


# ------------------------------------------------------------------ utilidades XML
def ln(el):
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""


def ch(el, nombre):
    if el is None:
        return None
    for c in el:
        if ln(c) == nombre:
            return c
    return None


def chs(el, nombre):
    return [c for c in el if ln(c) == nombre] if el is not None else []


def ruta(el, *nombres):
    for n in nombres:
        el = ch(el, n)
        if el is None:
            return None
    return el


def tx(el, *nombres):
    e = ruta(el, *nombres) if nombres else el
    return (e.text or "").strip() if e is not None and e.text else ""


# --------------------------------------------------------------- utilidades varias
def normaliza(texto):
    texto = unicodedata.normalize("NFKD", texto or "")
    return "".join(c for c in texto if not unicodedata.combining(c)).lower().strip()


def a_float(valor):
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def num(v):
    """30.0 -> '30'; 12.5 -> '12,5'"""
    if v is None:
        return ""
    return str(int(v)) if float(v).is_integer() else f"{v}".replace(".", ",")


def eur(v):
    if v is None:
        return "—"
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} €"


_TS = re.compile(r"(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(Z|[+-]\d{2}:\d{2})?")


def a_utc(s):
    """Marca de tiempo ATOM -> 'AAAA-MM-DDTHH:MM:SSZ' en UTC (comparable como texto)."""
    m = _TS.match(s or "")
    if not m:
        return ""
    a, mes, d, h, mi, se, tz = m.groups()
    dt = datetime(int(a), int(mes), int(d), int(h), int(mi), int(se))
    if tz and tz != "Z":
        signo = 1 if tz[0] == "+" else -1
        dt -= signo * timedelta(hours=int(tz[1:3]), minutes=int(tz[4:6]))
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def limite_dt(fecha):
    """'AAAA-MM-DD HH:MM' (hora local) -> datetime; sin hora = final del día."""
    if not fecha:
        return None
    try:
        if len(fecha) >= 16:
            return datetime.strptime(fecha[:16], "%Y-%m-%d %H:%M")
        return datetime.strptime(fecha[:10], "%Y-%m-%d") + timedelta(hours=23, minutes=59)
    except ValueError:
        return None


def descarga(url, destino=None, reintentos=3, timeout=120):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (licitaciones_obras.py)"})
    for intento in range(1, reintentos + 1):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                if destino is None:
                    return resp.read()
                total = int(resp.headers.get("Content-Length") or 0)
                leido = 0
                with open(destino + ".part", "wb") as f:
                    while True:
                        bloque = resp.read(1 << 20)
                        if not bloque:
                            break
                        f.write(bloque)
                        leido += len(bloque)
                        if total:
                            print(f"\r    {leido / 1e6:,.0f} / {total / 1e6:,.0f} MB", end="", flush=True)
                print()
                os.replace(destino + ".part", destino)
                return destino
        except Exception as exc:  # noqa: BLE001
            if intento == reintentos:
                raise
            print(f"  Aviso: fallo de descarga ({exc}); reintento {intento}/{reintentos - 1}...",
                  file=sys.stderr)
            time.sleep(3 * intento)


# --------------------------------------------------------------------- extracción
def extrae(entry):
    """Datos de una entrada ATOM, o None si no es una licitación."""
    cfs = ch(entry, "ContractFolderStatus")
    if cfs is None:
        return None
    pp = ch(cfs, "ProcurementProject")
    tp = ch(cfs, "TenderingProcess")
    tt = ch(cfs, "TenderingTerms")
    lotes = chs(cfs, "ProcurementProjectLot")

    cpvs = []
    for proyecto in [pp] + [ch(l, "ProcurementProject") for l in lotes]:
        for rc in chs(proyecto, "RequiredCommodityClassification"):
            c = tx(rc, "ItemClassificationCode")
            if c and c not in cpvs:
                cpvs.append(c)

    loc = ch(pp, "RealizedLocation")
    fecha = tx(tp, "TenderSubmissionDeadlinePeriod", "EndDate")
    hora = tx(tp, "TenderSubmissionDeadlinePeriod", "EndTime")
    if fecha and hora:
        fecha = f"{fecha} {hora[:5]}"

    dm = ruta(pp, "PlannedPeriod", "DurationMeasure")
    plazo = ""
    if dm is not None and (dm.text or "").strip():
        unidad = dm.get("unitCode", "")
        plazo = f"{num(a_float(dm.text.strip()))} {UNIDADES.get(unidad, unidad)}".strip()

    clasif = []
    for q in chs(tt, "TendererQualificationRequest"):
        for esquema in chs(q, "RequiredBusinessClassificationScheme"):
            for cat in chs(esquema, "ClassificationCategory"):
                v = tx(cat, "CodeValue")
                if v and v not in clasif:
                    clasif.append(v)

    criterios, peso_total, peso_jv = [], 0.0, 0.0
    lista_criterios = chs(ch(tt, "AwardingTerms"), "AwardingCriteria")
    if not lista_criterios and lotes:  # con lotes, los criterios van dentro de cada lote
        lista_criterios = chs(ruta(lotes[0], "TenderingTerms", "AwardingTerms"), "AwardingCriteria")
        if lista_criterios:
            criterios.append("[lote 1]")
    for a in lista_criterios:
        tipo = tx(a, "AwardingCriteriaTypeCode")
        peso = a_float(tx(a, "WeightNumeric"))
        etiqueta = (tx(a, "Description") or "¿?").replace("\n", " ")
        criterios.append(f"{etiqueta} {num(peso)}{' (JV)' if tipo == 'SUBJ' else ''}".strip())
        if peso:
            peso_total += peso
            if tipo == "SUBJ":
                peso_jv += peso

    garantias = []
    for g in chs(tt, "RequiredFinancialGuarantee"):
        codigo = tx(g, "GuaranteeTypeCode")
        tasa = tx(g, "AmountRate")
        if tasa:
            garantias.append(f"{GARANTIAS.get(codigo, codigo)} {num(a_float(tasa))}%")

    revision = " ".join(tx(tt, "PriceRevisionFormulaDescription").split())
    if normaliza(revision) in ("n/a", "na", "no", "-", "no procede", "no aplica"):
        revision = "No"

    link = entry.find(f"{ATOM}link")
    return {
        "id": tx(entry, "id"),
        "actualizado": a_utc(tx(entry, "updated")),
        "estado": tx(cfs, "ContractFolderStatusCode"),
        "tipo": tx(pp, "TypeCode"),
        "expediente": tx(cfs, "ContractFolderID"),
        "organo": tx(cfs, "LocatedContractingParty", "Party", "PartyName", "Name"),
        "objeto": " ".join((tx(pp, "Name") or tx(entry, "title")).split()),
        "cpv": cpvs,
        "importe": a_float(tx(pp, "BudgetAmount", "TaxExclusiveAmount")),
        "valor_estimado": a_float(tx(pp, "BudgetAmount", "EstimatedOverallContractAmount")),
        "procedimiento": tx(tp, "ProcedureCode"),
        "fecha_limite": fecha,
        "nuts": tx(loc, "CountrySubentityCode").upper(),
        "provincia_txt": tx(loc, "CountrySubentity"),
        "municipio": tx(loc, "Address", "CityName"),
        "plazo_ejecucion": plazo,
        "clasificacion": clasif,
        "criterios": criterios,
        "pct_jv": round(100 * peso_jv / peso_total, 1) if peso_total else None,
        "garantias": garantias,
        "revision_precios": revision,
        "lotes": len(lotes),
        "url_pcap": tx(cfs, "LegalDocumentReference", "Attachment", "ExternalReference", "URI"),
        "url_ppt": tx(cfs, "TechnicalDocumentReference", "Attachment", "ExternalReference", "URI"),
        "url": (link.get("href", "") if link is not None else "") or tx(entry, "id"),
    }


# -------------------------------------------------------------- ámbito geográfico
PROVINCIAS_TEXTO = {  # respaldo cuando una entrada no trae código NUTS
    "ES52": ["alicante", "alacant", "castellon", "castello", "valencia"],
    "ES61": ["almeria", "cadiz", "cordoba", "granada", "huelva", "jaen", "malaga", "sevilla"],
    "ES62": ["murcia"],
}


def en_ambito(reg, nuts_list, lugares):
    if not nuts_list and not lugares:
        return True
    cod = reg["nuts"]
    prov = normaliza(reg["provincia_txt"])
    for pref in nuts_list:
        if cod:
            if cod.startswith(pref):
                return True
        elif any(p in prov for p in PROVINCIAS_TEXTO.get(pref[:4], [])):
            return True
    texto = " | ".join(normaliza(x) for x in (reg["provincia_txt"], reg["nuts"], reg["municipio"]))
    return any(normaliza(l) in texto for l in lugares)


def es_de_obras(reg):
    """Para guardar en la base: tipo de contrato obras o algún CPV 45."""
    return reg["tipo"] == "3" or any(c.startswith("45") for c in reg["cpv"])


# ------------------------------------------------------------------- base local
class Base:
    def __init__(self, ruta_fichero):
        self.ruta = ruta_fichero
        self.datos = {"ultima_sincronizacion": "", "licitaciones": {}}
        if os.path.exists(ruta_fichero):
            with open(ruta_fichero, encoding="utf-8") as f:
                self.datos = json.load(f)
        self.nuevas, self.cambiadas = set(), set()
        self.hoy = date.today().isoformat()

    @property
    def lic(self):
        return self.datos["licitaciones"]

    def integra(self, reg):
        clave = reg["id"] or reg["expediente"]
        previo = self.lic.get(clave)
        if previo is None:
            reg["primera_deteccion"] = self.hoy
            reg["cambios"] = []
            self.lic[clave] = reg
            self.nuevas.add(clave)
            self._aplica_baja(clave)
            return
        if reg["actualizado"] <= previo.get("actualizado", ""):
            return  # ya tenemos una versión igual o más reciente
        dif = []
        if previo.get("estado") != reg["estado"]:
            dif.append(f"estado {previo.get('estado') or '—'} → {reg['estado'] or '—'}")
        if previo.get("fecha_limite") != reg["fecha_limite"]:
            dif.append(f"plazo {previo.get('fecha_limite') or '—'} → {reg['fecha_limite'] or '—'}")
        if previo.get("importe") != reg["importe"]:
            dif.append(f"importe {eur(previo.get('importe'))} → {eur(reg['importe'])}")
        if clave in self.nuevas:
            dif = []  # detectada en esta misma ejecución: no se anotan versiones intermedias
        reg["primera_deteccion"] = previo.get("primera_deteccion", self.hoy)
        reg["cambios"] = previo.get("cambios", []) + ([f"{self.hoy}: {'; '.join(dif)}"] if dif else [])
        self.lic[clave] = reg
        if dif and clave not in self.nuevas:
            self.cambiadas.add(clave)
        self._aplica_baja(clave)

    @property
    def bajas(self):
        return self.datos.setdefault("bajas", {})

    def borrada(self, ref, motivo, cuando):
        """Registra una baja (anulación/retirada). Se aplica aunque la licitación se lea después."""
        if not ref:
            return
        previa = self.bajas.get(ref)
        if previa is None or cuando > previa[0]:
            self.bajas[ref] = [cuando, motivo or "RETIRADA"]
        self._aplica_baja(ref)

    def _aplica_baja(self, ref):
        baja, r = self.bajas.get(ref), self.lic.get(ref)
        if not baja or not r:
            return
        cuando, motivo = baja
        if r.get("actualizado", "") <= cuando and r.get("estado") != motivo:
            r["cambios"] = r.get("cambios", []) + [f"{self.hoy}: {motivo.lower()} en la Plataforma"]
            r["estado"] = motivo
            if ref not in self.nuevas:
                self.cambiadas.add(ref)

    def guarda(self):
        carpeta = os.path.dirname(self.ruta)
        if carpeta:
            os.makedirs(carpeta, exist_ok=True)
        tmp = self.ruta + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.datos, f, ensure_ascii=False, indent=0)
        os.replace(tmp, self.ruta)


# --------------------------------------------------------------- procesado ATOM
class Lector:
    def __init__(self, base, nuts_list, lugares):
        self.base, self.nuts_list, self.lugares = base, nuts_list, lugares
        self.entradas = 0
        self.mas_antigua = ""
        self.mas_reciente = ""

    def procesa(self, contenido_xml):
        """Procesa una página ATOM. Devuelve (enlace_next, fecha_más_antigua_de_la_página)."""
        raiz = ET.fromstring(contenido_xml)
        antigua = ""
        for el in raiz:
            nombre = ln(el)
            if nombre == "entry":
                self.entradas += 1
                reg = extrae(el)
                if reg is None:
                    continue
                act = reg["actualizado"]
                if act:
                    antigua = act if not antigua or act < antigua else antigua
                    self.mas_antigua = act if not self.mas_antigua or act < self.mas_antigua else self.mas_antigua
                    self.mas_reciente = max(self.mas_reciente, act)
                if es_de_obras(reg) and en_ambito(reg, self.nuts_list, self.lugares):
                    self.base.integra(reg)
            elif nombre == "deleted-entry":
                comentario = ch(el, "comment")
                self.base.borrada(el.get("ref", ""),
                                  comentario.get("type", "") if comentario is not None else "",
                                  a_utc(el.get("when", "")))
        siguiente = None
        for l in raiz.findall(f"{ATOM}link"):
            if l.get("rel") == "next":
                siguiente = l.get("href")
        return siguiente, antigua


def lee_feed(lector, max_paginas, hasta, url=FEED_URL):
    """Lee el feed en vivo hasta max_paginas o hasta llegar a lo ya sincronizado."""
    paginas = 0
    while url and paginas < max_paginas:
        paginas += 1
        print(f"  Página {paginas}...", end=" ", flush=True)
        try:
            siguiente, antigua = lector.procesa(descarga(url))
        except Exception as exc:  # noqa: BLE001
            print(f"\n  ERROR al leer la página {paginas}: {exc}. Se detiene la lectura.", file=sys.stderr)
            return paginas - 1
        print(f"hasta {antigua[:16].replace('T', ' ')} UTC")
        if hasta and antigua and antigua < hasta:
            print("  Alcanzado el punto de la sincronización anterior.")
            break
        url = urljoin(url, siguiente) if siguiente else None
    return paginas


def lee_zip(lector, aaaamm, carpeta="zips"):
    os.makedirs(carpeta, exist_ok=True)
    destino = os.path.join(carpeta, f"placsp_{aaaamm}.zip")
    mes_actual = date.today().strftime("%Y%m")
    if not os.path.exists(destino) or aaaamm == mes_actual:  # el del mes en curso cambia cada día
        print(f"  Descargando fichero mensual {aaaamm}...")
        descarga(ZIP_URL.format(aaaamm), destino)
    with zipfile.ZipFile(destino) as z:
        nombres = [n for n in z.namelist() if n.lower().endswith(".atom")]
        print(f"  {aaaamm}: {len(nombres)} ficheros ATOM")
        for n in nombres:
            lector.procesa(z.read(n))


def meses_atras(n):
    d = date.today().replace(day=1)
    res = []
    for _ in range(n):
        res.append(d.strftime("%Y%m"))
        d = (d - timedelta(days=1)).replace(day=1)
    return sorted(res)


# -------------------------------------------------------------------- informe
def a_fila(clave, r, nuevas, ahora):
    lim = limite_dt(r["fecha_limite"])
    nuts = r["nuts"]
    lugar = ", ".join(x for x in (r["municipio"], r["provincia_txt"]) if x)
    lugar = f"{lugar} ({nuts})" if lugar and nuts else (lugar or nuts)
    coma = lambda v: f"{v:.2f}".replace(".", ",") if v is not None else ""
    comunidad = COMUNIDADES.get(nuts[:4], "")
    if not comunidad:  # sin código NUTS: se deduce del nombre de la provincia
        prov = normaliza(r["provincia_txt"])
        for cod, nombres in PROVINCIAS_TEXTO.items():
            if prov and any(n in prov for n in nombres):
                comunidad = COMUNIDADES[cod]
                break
    return {
        "nueva": "SÍ" if clave in nuevas else "",
        "expediente": r["expediente"],
        "organo_contratacion": r["organo"],
        "objeto": r["objeto"],
        "presupuesto_base_sin_iva": coma(r["importe"]),
        "valor_estimado": coma(r["valor_estimado"]),
        "fecha_limite_presentacion": r["fecha_limite"],
        "dias_restantes": (lim.date() - ahora.date()).days if lim else "",
        "comunidad": comunidad,
        "provincia": PROVINCIAS.get(nuts[:5], "" if len(nuts) == 4 else r["provincia_txt"]),
        "lugar_ejecucion": lugar,
        "procedimiento": PROCEDIMIENTOS.get(r["procedimiento"], r["procedimiento"]),
        "plazo_ejecucion": r["plazo_ejecucion"],
        "clasificacion_exigida": "; ".join(r["clasificacion"]),
        "criterios_adjudicacion": " | ".join(r["criterios"]),
        "pct_juicio_valor": num(r["pct_jv"]),
        "garantias": "; ".join(r["garantias"]),
        "revision_precios": r["revision_precios"],
        "num_lotes": r["lotes"] or "",
        "cpv": "; ".join(r["cpv"]),
        "url_pcap": r["url_pcap"],
        "url_ppt": r["url_ppt"],
        "url_licitacion": r["url"],
        "primera_deteccion": r.get("primera_deteccion", ""),
        "cambios": " | ".join(r.get("cambios", [])),
    }


def escribe_csv(filas, ruta_csv):
    carpeta = os.path.dirname(ruta_csv)
    if carpeta:
        os.makedirs(carpeta, exist_ok=True)
    try:
        f = open(ruta_csv, "w", newline="", encoding="utf-8-sig")
    except PermissionError:  # el CSV está abierto en Excel
        alt = ruta_csv.replace(".csv", f"_{datetime.now():%H%M%S}.csv")
        print(f"  Aviso: {ruta_csv} está abierto (¿en Excel?). Se guarda como {alt}")
        ruta_csv, f = alt, open(alt, "w", newline="", encoding="utf-8-sig")
    with f:
        w = csv.DictWriter(f, fieldnames=CAMPOS_CSV, delimiter=";")
        w.writeheader()
        w.writerows(filas)
    return ruta_csv


# ----------------------------------------------------------------------- main
def main():
    for flujo in (sys.stdout, sys.stderr):  # evita errores de codificación en Windows al redirigir
        try:
            flujo.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--nuts", default="ES52,ES61",
                    help="Prefijos NUTS separados por comas (por defecto ES52,ES61: C. Valenciana y Andalucía)")
    ap.add_argument("--lugar", default="", help="Además, texto de provincia o municipio, separado por comas")
    ap.add_argument("--paginas", type=int, default=30,
                    help="Máximo de páginas del feed en vivo (se detiene antes si llega a lo ya leído)")
    ap.add_argument("--meses", type=int, default=0,
                    help="Carga los N últimos ficheros ZIP mensuales (carga inicial; recomendado 2)")
    ap.add_argument("--zip", action="append", default=[], metavar="AAAAMM",
                    help="Carga un fichero mensual concreto; se puede repetir")
    ap.add_argument("--min", type=float, default=None, help="Presupuesto mínimo sin IVA para el CSV")
    ap.add_argument("--max", type=float, default=None, help="Presupuesto máximo sin IVA para el CSV")
    ap.add_argument("--incluir-mixtos", action="store_true",
                    help="Incluir en el CSV contratos que no son de tipo obras pero tienen algún CPV 45")
    ap.add_argument("--base", default=BASE_POR_DEFECTO, help="Fichero de la base local")
    ap.add_argument("--salida", default=SALIDA_POR_DEFECTO, help="CSV de salida")
    ap.add_argument("--url", default=FEED_URL, help=argparse.SUPPRESS)
    args = ap.parse_args()

    nuts_list = [x.strip().upper() for x in args.nuts.split(",") if x.strip()]
    lugares = [x.strip() for x in args.lugar.split(",") if x.strip()]
    base = Base(args.base)
    primera_vez = not base.lic
    lector = Lector(base, nuts_list, lugares)
    inicio = datetime.now()

    zips = sorted(set(args.zip + meses_atras(args.meses))) if (args.zip or args.meses) else []
    for z in zips:
        try:
            lee_zip(lector, z)
        except Exception as exc:  # noqa: BLE001
            print(f"  ERROR con el fichero mensual {z}: {exc}", file=sys.stderr)

    print("Leyendo el feed en vivo...")
    # margen de 1 día sobre la última sincronización para no perder nada
    hasta = ""
    if base.datos.get("ultima_sincronizacion"):
        ult = datetime.strptime(base.datos["ultima_sincronizacion"], "%Y-%m-%dT%H:%M:%SZ")
        hasta = (ult - timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%SZ")
    paginas = lee_feed(lector, args.paginas, hasta, args.url)

    if lector.mas_reciente:
        base.datos["ultima_sincronizacion"] = max(base.datos.get("ultima_sincronizacion", ""),
                                                  lector.mas_reciente)
    base.guarda()

    # Licitaciones abiertas
    ahora = datetime.now()
    abiertas = []
    for clave, r in base.lic.items():
        if r.get("estado") != "PUB":
            continue
        lim = limite_dt(r["fecha_limite"])
        if lim and lim < ahora:
            continue
        if not args.incluir_mixtos and r.get("tipo") != "3":
            continue
        if args.min is not None and (r["importe"] is None or r["importe"] < args.min):
            continue
        if args.max is not None and (r["importe"] is None or r["importe"] > args.max):
            continue
        abiertas.append((clave, r))
    abiertas.sort(key=lambda kr: kr[1]["fecha_limite"] or "9999")
    filas = [a_fila(k, r, base.nuevas, ahora) for k, r in abiertas]
    ruta_csv = escribe_csv(filas, args.salida)

    # Resumen
    nuevas_abiertas = [r for k, r in abiertas if k in base.nuevas]
    cambiadas_abiertas = [r for k, r in abiertas if k in base.cambiadas]
    total = sum(r["importe"] or 0 for _, r in abiertas)
    rango = ""
    if args.min is not None or args.max is not None:
        rango = f" entre {eur(args.min) if args.min is not None else '0 €'} y {eur(args.max) if args.max is not None else '∞'}"
    print("\n" + "=" * 78)
    print(f"Fuentes:  {len(zips)} fichero(s) mensual(es) + {paginas} página(s) del feed "
          f"({(datetime.now() - inicio).seconds} s)")
    print(f"Entradas leídas en total:     {lector.entradas}")
    if lector.mas_antigua:
        print(f"Periodo cubierto:             {lector.mas_antigua[:10]} a {lector.mas_reciente[:10]}")
    print(f"Licitaciones en la base:      {len(base.lic)}")
    if primera_vez:
        print("Primera ejecución: todas las licitaciones cuentan como nuevas.")
    print(f"\nEN PLAZO (obras{rango}):  {len(abiertas)}  |  Importe total: {eur(total)}")
    print(f"  Nuevas desde la última ejecución:  {len(nuevas_abiertas)}")
    print(f"  Con cambios (plazo, importe...):   {len(cambiadas_abiertas)}")
    print(f"CSV guardado en: {ruta_csv}")

    urgentes = [r for _, r in abiertas if limite_dt(r["fecha_limite"]) and
                (limite_dt(r["fecha_limite"]).date() - ahora.date()).days <= 7]
    if urgentes:
        print("\nVencen en 7 días o menos:")
        for r in urgentes:
            print(f"  {r['fecha_limite']}  {eur(r['importe']):>16}  {r['organo'][:40]} | {r['objeto'][:50]}")
    if not primera_vez and nuevas_abiertas:
        print("\nNuevas (por importe):")
        for r in sorted(nuevas_abiertas, key=lambda r: r["importe"] or 0, reverse=True)[:10]:
            print(f"  {eur(r['importe']):>16}  plazo {r['fecha_limite']}  {r['organo'][:40]} | {r['objeto'][:50]}")
    if cambiadas_abiertas:
        print("\nCambios detectados:")
        for r in cambiadas_abiertas:
            print(f"  {r['expediente']}: {r['cambios'][-1]}")
    print("=" * 78)


if __name__ == "__main__":
    main()
