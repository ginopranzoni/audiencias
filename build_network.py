#!/usr/bin/env python3
"""Build the GitHub Pages network data from audiencias_2004-2022_AFIPSDK_v6.xlsx.

No third-party Python packages are required. The .xlsx file is read as its native
Open XML ZIP package so the script can process the ~29 MB workbook with low memory.

Outputs:
  - nodes.json
  - edges.json
  - network_meta.json

Network definition:
  public node  = public organism/dependency receiving the audience
  private node = represented actor when present; otherwise the requesting actor
  edge weight  = number of unique audiences between organism and actor

Only edge pairs with >= MIN_TOTAL_EDGE_WEIGHT audiences over 2004-2022 are exported.
The web interface can apply a higher threshold and year/sector/type filters.
"""

import hashlib
import json
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path

XLSX = Path("audiencias_2004-2022_AFIPSDK_v6.xlsx")
OUT_DIR = Path(".")
MIN_TOTAL_EDGE_WEIGHT = 5
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Workbook columns used by the network.
COLS = {
    "A": "audiencia_id",
    "C": "fecha_audiencia",
    "L": "cargo_dependencia",
    "M": "root_dependencia",
    "N": "solic_id",
    "P": "solic_nombre",
    "S": "caracter",
    "T": "representado_id",
    "U": "representado_nombre",
    "W": "representado_persona_juridica",
    "AE": "estado",
    "AJ": "tipo_persona",
    "AL": "actividad_principal",
    "AM": "id_actividad_principal",
    "AN": "forma_juridica",
    "AQ": "categoria_actividad",
}


def col_letters(ref):
    return re.match(r"([A-Z]+)", ref).group(1)


def norm(value):
    value = (value or "").strip()
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = value.upper()
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def only_digits(value):
    return "".join(ch for ch in (value or "") if ch.isdigit())


def excel_year(value):
    try:
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).year
    except Exception:
        return None


def public_organism(row):
    """Use the top-level dependency when available; after 2016 fall back to the
    dependency embedded in obligado_cargo_dependencia.
    """
    root = (row.get("root_dependencia") or "").strip()
    if root:
        return root
    cargo_dep = (row.get("cargo_dependencia") or "").strip()
    if " - " in cargo_dep:
        return cargo_dep.split(" - ", 1)[1].strip()
    return cargo_dep


def actor_name(row):
    represented = (row.get("representado_nombre") or "").strip()
    applicant = (row.get("solic_nombre") or "").strip()
    return represented or applicant


def actor_id(row, name):
    """Create a stable, non-identifying node id.

    The raw document/tax identifier is never written to the JSON output.
    """
    raw = (
        (row.get("representado_id") or "").strip()
        if (row.get("representado_nombre") or "").strip()
        else (row.get("solic_id") or "").strip()
    )
    digits = only_digits(raw)
    if len(digits) == 11:
        identity = "doc11:" + digits
    elif raw and len(norm(raw)) >= 5:
        identity = "doc:" + norm(raw) + "|name:" + norm(name)
    else:
        identity = "name:" + norm(name)
    return "actor_" + hashlib.sha1(identity.encode("utf-8")).hexdigest()[:14]


def sector_from_activity(activity, activity_code="", fallback_category=""):
    """Compress descripcionActividadPrincipal into a small set of macro-sectors.

    AL (descripcionActividadPrincipal) is deliberately the first and main signal.
    The activity code and AQ are only fallbacks when the description is absent or
    not sufficiently informative.
    """
    t = norm(activity)

    text_rules = [
        ("Agro, silvicultura y pesca", [
            "CULTIVO", "AGRICULT", "GANADER", "CRIA DE", "PESCA", "SILVICULT", "SERVICIOS AGRICOL"
        ]),
        ("Minería, petróleo y gas", [
            "EXTRACCION DE PETROLEO", "EXTRACCION DE GAS", "PETROLEO CRUDO", "PETROLER",
            "MINER", "MINAS", "CANTERAS", "EXTRACCION DE MINERAL"
        ]),
        ("Energía, agua y saneamiento", [
            "ENERGIA ELECTR", "ELECTRIC", "DISTRIBUCION DE GAS", "GAS POR RED", "FABRICACION Y DISTRIBUCION DE GAS",
            "SUMINISTRO DE AGUA", "CAPTACION DE AGUA", "CLOACAS", "SANEAMIENTO", "RESIDUOS"
        ]),
        ("Construcción", ["CONSTRUCCION", "OBRAS DE INGENIERIA CIVIL", "OBRAS VIALES"]),
        ("Industria manufacturera", [
            "FABRICACION", "ELABORACION", "MANUFACT", "IMPRESION", "REFINACION", "REPARACION DE MAQUINARIA"
        ]),
        ("Comercio", ["COMERCIO", "VENTA AL POR MAYOR", "VENTA AL POR MENOR", "REPARACION DE VEHICULOS"]),
        ("Transporte y logística", [
            "TRANSPORTE", "ALMACENAMIENTO", "LOGIST", "SERVICIOS PORTUARIOS", "SERVICIOS AEROPORTUARIOS", "CORREO"
        ]),
        ("Alojamiento y gastronomía", ["ALOJAMIENTO", "HOTEL", "RESTAURANTE", "SERVICIOS DE COMIDA", "GASTRONOM"]),
        ("Información y comunicaciones", [
            "INFORMAT", "SOFTWARE", "TELECOMUNIC", "COMUNICACION", "TRANSMISION DE RADIO", "TELEVISION",
            "EDICION", "PROCESAMIENTO DE DATOS", "PORTALES WEB"
        ]),
        ("Finanzas, seguros e inmobiliario", [
            "FINANCI", "BANC", "SEGURO", "MERCADO DE VALORES", "INMOBILIAR"
        ]),
        ("Educación, salud y servicios sociales", [
            "ENSENAN", "EDUCACION", "SALUD", "MEDIC", "HOSPITAL", "SERVICIOS SOCIALES", "ODONTO"
        ]),
        ("Asociaciones, cultura y otros servicios", [
            "ASOCIACION", "CAMARA", "GREMIO", "SINDIC", "FEDERACION", "FUNDACION", "CULTURAL", "ARTIST",
            "DEPORT", "ESPARCIMIENTO", "SERVICIOS PERSONALES", "BIBLIOTEC", "ARCHIVO"
        ]),
        ("Servicios profesionales y empresariales", [
            "JURID", "CONTABIL", "AUDITOR", "ASESORAMIENTO", "DIRECCION Y GESTION EMPRESARIAL", "ARQUITECT",
            "INGENIER", "PUBLICIDAD", "INVESTIGACION", "CONSULTOR", "SERVICIOS EMPRESARIALES", "ALQUILER DE MAQUINARIA",
            "ACTIVIDADES ADMINISTRATIVAS", "SERVICIOS PROFESIONALES", "CIENTIF"
        ]),
        ("Administración pública", [
            "ADMINISTRACION PUBLICA", "ASUNTOS EXTERIORES", "DEFENSA", "SEGURIDAD SOCIAL OBLIGATORIA"
        ]),
    ]
    if t:
        for label, keywords in text_rules:
            if any(norm(keyword) in t for keyword in keywords):
                return label

    # Fallback for rows where AL is absent/uninformative: broad AQ labels.
    aq = norm(fallback_category)
    aq_map = {
        "AGRICULTURA GANADERIA CAZA SILVICULTURA Y PESCA": "Agro, silvicultura y pesca",
        "EXPLOTACION DE MINAS Y CANTERAS": "Minería, petróleo y gas",
        "INDUSTRIA MANUFACTURERA": "Industria manufacturera",
        "SUMINISTRO DE ELECTRICIDAD GAS VAPOR Y AIRE ACONDICIONADO": "Energía, agua y saneamiento",
        "SUMINISTRO DE AGUA CLOACAS GESTION DE RESIDUOS Y RECUPERACION DE MATERIALES Y SANEAMIENTO PUBLICO": "Energía, agua y saneamiento",
        "CONSTRUCCION": "Construcción",
        "COMERCIO AL POR MAYOR Y AL POR MENOR REPARACION DE VEHICULOS AUTOMOTORES Y MOTOCICLETAS": "Comercio",
        "SERVICIO DE TRANSPORTE Y ALMACENAMIENTO": "Transporte y logística",
        "SERVICIOS DE ALOJAMIENTO Y SERVICIOS DE COMIDA": "Alojamiento y gastronomía",
        "INFORMACION Y COMUNICACIONES": "Información y comunicaciones",
        "INTERMEDIACION FINANCIERA Y SERVICIOS DE SEGUROS": "Finanzas, seguros e inmobiliario",
        "SERVICIOS INMOBILIARIOS": "Finanzas, seguros e inmobiliario",
        "SERVICIOS PROFESIONALES CIENTIFICOS Y TECNICOS": "Servicios profesionales y empresariales",
        "ACTIVIDADES ADMINISTRATIVAS Y SERVICIOS DE APOYO": "Servicios profesionales y empresariales",
        "ENSENANZA": "Educación, salud y servicios sociales",
        "SALUD HUMANA Y SERVICIOS SOCIALES": "Educación, salud y servicios sociales",
        "SERVICIOS DE ASOCIACIONES Y SERVICIOS PERSONALES": "Asociaciones, cultura y otros servicios",
        "SERVICIOS ARTISTICOS CULTURALES DEPORTIVOS Y DE ESPARCIMIENTO": "Asociaciones, cultura y otros servicios",
        "ADMINISTRACION PUBLICA DEFENSA Y SEGURIDAD SOCIAL OBLIGATORIA": "Administración pública",
    }
    if aq in aq_map:
        return aq_map[aq]

    # Last fallback: section ranges from the activity code when possible.
    d = only_digits(activity_code)
    if len(d) >= 2:
        try:
            p = int(d[:2])
            if 1 <= p <= 3: return "Agro, silvicultura y pesca"
            if 5 <= p <= 9: return "Minería, petróleo y gas"
            if 10 <= p <= 33: return "Industria manufacturera"
            if 35 <= p <= 39: return "Energía, agua y saneamiento"
            if 41 <= p <= 43: return "Construcción"
            if 45 <= p <= 47: return "Comercio"
            if 49 <= p <= 53: return "Transporte y logística"
            if 55 <= p <= 56: return "Alojamiento y gastronomía"
            if 58 <= p <= 63: return "Información y comunicaciones"
            if 64 <= p <= 68: return "Finanzas, seguros e inmobiliario"
            if 69 <= p <= 82: return "Servicios profesionales y empresariales"
            if p == 84: return "Administración pública"
            if 85 <= p <= 88: return "Educación, salud y servicios sociales"
            if 90 <= p <= 96: return "Asociaciones, cultura y otros servicios"
        except ValueError:
            pass
    return "Sin clasificar"


def obvious_public_name(name):
    n = norm(name)
    patterns = [
        "PRESIDENCIA DE LA NACION", "MINISTERIO DE ", "SECRETARIA DE ", "SUBSECRETARIA DE ",
        "MUNICIPALIDAD DE ", "GOBIERNO DE LA PROVINCIA", "GOBIERNO PROVINCIAL", "GOBIERNO NACIONAL",
        "DIRECCION NACIONAL DE ", "ADMINISTRACION NACIONAL DE ", "SENADO DE LA NACION",
        "CAMARA DE DIPUTADOS", "BANCO CENTRAL DE LA REPUBLICA ARGENTINA"
    ]
    return any(p in n for p in patterns)


def actor_type(row, name):
    ch = norm(row.get("caracter"))
    forma = norm(row.get("forma_juridica"))
    activity = norm(row.get("actividad_principal"))
    n = norm(name)
    tipo = norm(row.get("tipo_persona"))

    if (
        "ORGANISMO ESTATAL" in ch
        or "ORGAN PUBLICO" in forma
        or "DIR ADM ESTATAL" in forma
        or "EMP DEL ESTADO" in forma
        or "FIDEICOMISO PUBLICO" in forma
        or obvious_public_name(name)
    ):
        return "Actor público"

    # Specific organizational subtypes first.
    if "SINDIC" in n or "SINDIC" in activity or "GREMIO" in n or "UNION DE TRABAJ" in n:
        return "Sindicato / gremio"
    if "CAMARA" in n or "CONFEDERACION" in n or "FEDERACION" in n or "ASOCIACION EMPRES" in n:
        return "Cámara / asociación empresaria"
    if forma == "FUNDACION" or "FUNDACION" in n or n.startswith("ONG ") or " ONG " in (" " + n + " "):
        return "Fundación / ONG"
    if forma in {"COOPERATIVA", "MUTUAL"} or "COOPERATIVA" in n or "MUTUAL" in n:
        return "Cooperativa / mutual"

    # Character in the audience registry has priority over the AFIP person type.
    if "GRUPO DE PERSONAS" in ch:
        return "Grupo de personas"
    if "REPRESENTANTE DE PERSONA JURIDICA" in ch or row.get("representado_persona_juridica") == "1":
        if forma == "ASOCIACION" or "ASOCIACION" in n:
            return "Asociación / entidad civil"
        return "Empresa / entidad privada"
    if "REPRESENTANTE DE PERSONA FISICA" in ch or "PARTICULAR INTERESADO" in ch:
        return "Persona física"

    # AFIP data as fallback.
    if tipo == "JURIDICA":
        if forma == "ASOCIACION" or "ASOCIACION" in n:
            return "Asociación / entidad civil"
        return "Empresa / entidad privada"
    if tipo == "FISICA":
        return "Persona física"
    return "Otro actor privado"


def read_shared_strings(zf):
    strings = []
    with zf.open("xl/sharedStrings.xml") as fh:
        for _, element in ET.iterparse(fh, events=("end",)):
            if element.tag == NS + "si":
                strings.append("".join((t.text or "") for t in element.iter(NS + "t")))
                element.clear()
    return strings


def iter_rows(zf, shared_strings):
    with zf.open("xl/worksheets/sheet1.xml") as fh:
        for _, element in ET.iterparse(fh, events=("end",)):
            if element.tag != NS + "row":
                continue
            if element.attrib.get("r") == "1":
                element.clear()
                continue
            row = {}
            for cell in element.findall(NS + "c"):
                c = col_letters(cell.attrib.get("r", ""))
                if c not in COLS:
                    continue
                value_node = cell.find(NS + "v")
                raw = "" if value_node is None else (value_node.text or "")
                typ = cell.attrib.get("t")
                if raw and typ == "s":
                    try:
                        value = shared_strings[int(raw)]
                    except Exception:
                        value = raw
                elif typ == "inlineStr":
                    value = "".join((t.text or "") for t in cell.iter(NS + "t"))
                else:
                    value = raw
                row[COLS[c]] = value.strip() if isinstance(value, str) else value
            yield row
            element.clear()


def main():
    if not XLSX.exists():
        raise SystemExit(f"No encuentro {XLSX}. Colocá este script junto al Excel v6.")

    edges_by_year = defaultdict(Counter)
    actor_stats = {}
    edge_seen = set()
    excluded = Counter()
    included_audiences = set()

    with zipfile.ZipFile(XLSX) as zf:
        shared_strings = read_shared_strings(zf)
        for row in iter_rows(zf, shared_strings):
            state = (row.get("estado") or "").strip()
            # 2017-2022 have blank status in v6, so blanks are retained.
            if state in {"No Realizada", "Derivada", "Reservada"}:
                excluded["estado_no_reunion"] += 1
                continue

            organism = public_organism(row)
            if not organism:
                excluded["sin_organismo"] += 1
                continue

            year = excel_year(row.get("fecha_audiencia"))
            if year is None or not (2004 <= year <= 2022):
                excluded["sin_anio"] += 1
                continue

            name = actor_name(row)
            if not name:
                excluded["sin_actor"] += 1
                continue

            a_type = actor_type(row, name)
            if a_type == "Actor público":
                excluded["actor_publico"] += 1
                continue

            a_id = actor_id(row, name)
            audience_id = (row.get("audiencia_id") or "").strip()
            unique_edge = (audience_id, organism, a_id)
            if unique_edge in edge_seen:
                excluded["duplicado_audiencia_actor"] += 1
                continue
            edge_seen.add(unique_edge)

            sector = sector_from_activity(
                row.get("actividad_principal", ""),
                row.get("id_actividad_principal", ""),
                row.get("categoria_actividad", ""),
            )

            edges_by_year[(organism, a_id)][year] += 1
            included_audiences.add(audience_id)

            stats = actor_stats.setdefault(
                a_id,
                {
                    "names": Counter(),
                    "types": Counter(),
                    "sectors": Counter(),
                    "activities": Counter(),
                    "forms": Counter(),
                    "years": Counter(),
                    "audiences": 0,
                },
            )
            stats["names"][name] += 1
            stats["types"][a_type] += 1
            stats["sectors"][sector] += 1
            if row.get("actividad_principal"):
                stats["activities"][row["actividad_principal"]] += 1
            if row.get("forma_juridica"):
                stats["forms"][row["forma_juridica"]] += 1
            stats["years"][year] += 1
            stats["audiences"] += 1

    def dominant(counter, default=""):
        return counter.most_common(1)[0][0] if counter else default

    actor_meta = {}
    for a_id, s in actor_stats.items():
        actor_meta[a_id] = {
            "label": dominant(s["names"]),
            "actorType": dominant(s["types"], "Otro actor privado"),
            "sector": dominant(s["sectors"], "Sin clasificar"),
            "activity": dominant(s["activities"]),
            "formaJuridica": dominant(s["forms"]),
            "totalAudiences": s["audiences"],
            "firstYear": min(s["years"]) if s["years"] else None,
            "lastYear": max(s["years"]) if s["years"] else None,
        }

    # Export only relationships that can ever be visible with the web's minimum threshold.
    kept_edges = []
    kept_actor_ids = set()
    kept_organisms = set()
    for (organism, a_id), years in edges_by_year.items():
        total = sum(years.values())
        if total < MIN_TOTAL_EDGE_WEIGHT:
            continue
        kept_actor_ids.add(a_id)
        kept_organisms.add(organism)
        kept_edges.append(
            {
                "data": {
                    "id": "edge_" + hashlib.sha1((organism + "|" + a_id).encode("utf-8")).hexdigest()[:14],
                    "source": "org_" + hashlib.sha1(organism.encode("utf-8")).hexdigest()[:14],
                    "target": a_id,
                    "weight": total,
                    "byYear": {str(y): n for y, n in sorted(years.items())},
                }
            }
        )

    nodes = []
    for organism in sorted(kept_organisms):
        nodes.append(
            {
                "data": {
                    "id": "org_" + hashlib.sha1(organism.encode("utf-8")).hexdigest()[:14],
                    "label": organism,
                    "side": "public",
                    "actorType": "Organismo público",
                    "sector": "Sector público",
                }
            }
        )

    for a_id in kept_actor_ids:
        meta = actor_meta[a_id]
        nodes.append(
            {
                "data": {
                    "id": a_id,
                    "label": meta["label"],
                    "side": "private",
                    "actorType": meta["actorType"],
                    "sector": meta["sector"],
                    "activity": meta["activity"],
                    "formaJuridica": meta["formaJuridica"],
                    "totalAudiences": meta["totalAudiences"],
                    "firstYear": meta["firstYear"],
                    "lastYear": meta["lastYear"],
                }
            }
        )

    # Deterministic order makes diffs in GitHub easier to inspect.
    nodes.sort(key=lambda x: (x["data"].get("side", ""), norm(x["data"].get("label", ""))))
    kept_edges.sort(key=lambda x: (-x["data"]["weight"], x["data"]["source"], x["data"]["target"]))

    # Useful UI lists and diagnostics.
    sectors = sorted({n["data"].get("sector") for n in nodes if n["data"].get("side") == "private"})
    actor_types = sorted({n["data"].get("actorType") for n in nodes if n["data"].get("side") == "private"})
    edge_thresholds = {}
    for threshold in (5, 8, 10, 15, 20, 30, 50):
        visible = [e for e in kept_edges if e["data"]["weight"] >= threshold]
        visible_node_ids = {e["data"]["source"] for e in visible} | {e["data"]["target"] for e in visible}
        edge_thresholds[str(threshold)] = {
            "edges": len(visible),
            "nodes": len(visible_node_ids),
            "audiences": sum(e["data"]["weight"] for e in visible),
        }

    meta = {
        "sourceFile": XLSX.name,
        "years": list(range(2004, 2023)),
        "exportMinEdgeWeight": MIN_TOTAL_EDGE_WEIGHT,
        "defaultAllYearsThreshold": 15,
        "defaultSingleYearThreshold": 5,
        "nodes": len(nodes),
        "actors": len(kept_actor_ids),
        "organisms": len(kept_organisms),
        "edges": len(kept_edges),
        "includedUniqueAudiencesBeforeEdgeThreshold": len(included_audiences),
        "sectors": sectors,
        "actorTypes": actor_types,
        "thresholdSummary": edge_thresholds,
        "excludedRows": dict(excluded),
        "methodology": {
            "actor": "representado_nombre cuando existe; en caso contrario solic_nombre_completo",
            "organism": "root_dependencia_descripcion; si falta, dependencia extraída de obligado_cargo_dependencia",
            "sector": "macro-sector derivado principalmente de descripcionActividadPrincipal (AL); AQ/código sólo como respaldo",
            "states": "se excluyen Derivada, No Realizada y Reservada; los estados vacíos se conservan porque v6 no informa estado en 2017-2022",
            "publicActors": "se excluyen registros identificados como organismos/actores públicos para concentrar la red en relaciones Estado-privados",
        },
    }

    (OUT_DIR / "nodes.json").write_text(json.dumps(nodes, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    (OUT_DIR / "edges.json").write_text(json.dumps(kept_edges, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    (OUT_DIR / "network_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
