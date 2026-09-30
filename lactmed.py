import os
import tarfile
from datetime import datetime
from urllib.request import urlopen
from urllib.error import URLError, HTTPError
import xml.etree.ElementTree as ET

TARBALL_URL = "https://ftp.ncbi.nlm.nih.gov/pub/litarch/90/6c/lactmed_NBK501922.tar.gz"
TARBALL_NAME = "lactmed_NBK501922.tar.gz"


def _istag(el, suffix):
    return isinstance(el.tag, str) and el.tag.endswith(suffix)


def _find_child(el, suffix):
    for c in el:
        if _istag(c, suffix):
            return c
    return None


def _texto(el):
    return "".join(el.itertext()).strip() if el is not None else ""


def _sec_cuerpo(sec):
    """Texto directo de una sección, sin el título ni subsecciones anidadas."""
    partes = []
    for c in sec:
        if not _istag(c, "title") and "sec" not in c.tag.lower():
            texto = "".join(c.itertext()).strip()
            if texto:
                partes.append(texto)
    return " ".join(partes).strip()


def parse_nxml(texto):
    """Parsea un registro LactMed en formato BITS (.nxml). Devuelve dict o None."""
    try:
        root = ET.fromstring(texto)
    except ET.ParseError:
        return None
    book_part = _find_child(root, "book-part")
    if book_part is None:
        return None
    meta = _find_child(book_part, "book-part-meta")
    drug = ""
    aliases = []
    if meta is not None:
        title_group = _find_child(meta, "title-group")
        if title_group is not None:
            drug = _texto(_find_child(title_group, "title"))
        kwd_group = _find_child(meta, "kwd-group")
        if kwd_group is not None:
            for k in kwd_group:
                if _istag(k, "kwd"):
                    a = " ".join(k.itertext()).strip()
                    if a:
                        aliases.append(a)
    if not drug:
        return None
    body = _find_child(book_part, "body")
    secciones = {}
    if body is not None:
        for sec in body.iter():
            if not _istag(sec, "sec"):
                continue
            t = ""
            for c in sec:
                if _istag(c, "title"):
                    t = "".join(c.itertext()).strip()
                    break
            if t:
                secciones[t.strip().lower()] = _sec_cuerpo(sec)
    summary = secciones.get("summary of use during lactation", "")
    efectos_infante = secciones.get("effects in breastfed infants", "")
    efectos_lactancia = secciones.get("effects on lactation and breastmilk", "")
    consideration = "\n\n".join([x for x in (efectos_infante, efectos_lactancia) if x])
    alternatives = secciones.get("alternate drugs to consider", "")
    return {
        "drug_name": drug,
        "aliases": " | ".join(aliases),
        "summary": summary,
        "consideration": consideration,
        "alternatives": alternatives,
    }


def descargar_tarball(dest_path):
    if os.path.exists(dest_path):
        return os.path.getsize(dest_path)
    try:
        with urlopen(TARBALL_URL, timeout=120) as response, open(dest_path, "wb") as out:
            out.write(response.read())
        return os.path.getsize(dest_path)
    except (URLError, HTTPError) as e:
        raise RuntimeError(f"Error al descargar {TARBALL_URL}: {e}")


def _registros_del_tarball(tar_path):
    with tarfile.open(tar_path, "r:gz") as tf:
        for mi in tf:
            if mi.isfile() and mi.name.lower().endswith(".nxml"):
                f = tf.extractfile(mi)
                texto = f.read().decode("utf-8", errors="replace")
                rec = parse_nxml(texto)
                if rec:
                    yield rec


def importar_registros(conn, registros):
    conn.execute("DELETE FROM lactmed")
    n = 0
    for rec in registros:
        conn.execute(
            "INSERT INTO lactmed (drug_name, aliases, summary, consideration, alternatives) "
            "VALUES (?,?,?,?,?)",
            (rec["drug_name"], rec["aliases"], rec["summary"], rec["consideration"], rec["alternatives"]),
        )
        n += 1
    conn.commit()
    return n


def descargar_e_importar(conn, local_dir):
    os.makedirs(local_dir, exist_ok=True)
    tar_path = os.path.join(local_dir, TARBALL_NAME)
    descargar_tarball(tar_path)
    n = importar_registros(conn, _registros_del_tarball(tar_path))
    conn.execute("DELETE FROM lactmed_meta WHERE id = 1")
    conn.execute(
        "INSERT INTO lactmed_meta (id, fuente, fecha) VALUES (1,?,?)",
        (TARBALL_NAME, datetime.now().isoformat()),
    )
    conn.commit()
    return n