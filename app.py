from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

BASELINE_CENTER = "San Jose"
# Ruta relativa a la raiz del repo; funciona en local y en Streamlit Cloud
_HERE = Path(__file__).parent
DEFAULT_EXCEL = str(_HERE / "docs" / "Template V4b - San José (6).xlsx")


def norm(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if len(text) >= 2 and text[0] == "'" and text[-1] == "'":
        text = text[1:-1].strip()
    return text


def compact_openpyxl_sheet(file_path: Path, sheet_name: str, empty_streak_limit: int = 2500) -> pd.DataFrame:
    """Read sheet rows until a long empty streak appears.

    Several sheets in this workbook have formatted ranges up to ~1M rows.
    This avoids loading those empty tails while preserving real data rows.
    """
    wb = load_workbook(file_path, data_only=True, read_only=True)
    ws = wb[sheet_name]

    rows: list[list[str]] = []
    empty_streak = 0
    max_seen_cols = 0

    for row in ws.iter_rows(values_only=True):
        vals = [norm(v) for v in row]
        if any(vals):
            rows.append(vals)
            max_seen_cols = max(max_seen_cols, len(vals))
            empty_streak = 0
        else:
            if rows:
                empty_streak += 1
                if empty_streak >= empty_streak_limit:
                    break

    wb.close()

    if not rows:
        return pd.DataFrame()

    rows = [r + [""] * (max_seen_cols - len(r)) for r in rows]
    header = rows[0]
    data = rows[1:]

    safe_header: list[str] = []
    seen: dict[str, int] = {}
    for i, col in enumerate(header):
        col_name = norm(col) or f"col_{i+1}"
        if col_name in seen:
            seen[col_name] += 1
            col_name = f"{col_name}_{seen[col_name]}"
        else:
            seen[col_name] = 1
        safe_header.append(col_name)

    return pd.DataFrame(data, columns=safe_header)


@st.cache_data(show_spinner=False)
def load_data(file_path: Path) -> dict[str, pd.DataFrame]:
    convenios = compact_openpyxl_sheet(file_path, "Convenios")
    convenios_planes = compact_openpyxl_sheet(file_path, "Convenios_Planes")
    homologaciones = compact_openpyxl_sheet(file_path, "Homologación")
    prestaciones_catalogos = compact_openpyxl_sheet(file_path, "PrestacionesCatalogos")
    catalogos_convenio = compact_openpyxl_sheet(file_path, "CatalogosConvenio")
    modulo_prestacion = compact_openpyxl_sheet(file_path, "ModuloPrestacion")
    modulos_reglas = compact_openpyxl_sheet(file_path, "ModulosReglas")

    return {
        "convenios": convenios,
        "convenios_planes": convenios_planes,
        "homologaciones": homologaciones,
        "prestaciones_catalogos": prestaciones_catalogos,
        "catalogos_convenio": catalogos_convenio,
        "modulo_prestacion": modulo_prestacion,
        "modulos_reglas": modulos_reglas,
    }


def merge_homologaciones_with_modulo(
    homologaciones: pd.DataFrame,
    prestaciones_catalogos: pd.DataFrame,
) -> pd.DataFrame:
    h = homologaciones.copy()
    p = prestaciones_catalogos.copy()

    needed_h = [
        "Catalogo",
        "codigo_referencia",
        "PrestacionReferencia",
        "Codigo_Catalogo",
        "PrestacionCatalogo",
        "Tipo Homologacion",
        "ESTADO",
        "fecha inicio vigencia",
    ]
    for col in needed_h:
        if col not in h.columns:
            h[col] = ""

    needed_p = ["Catalogo", "Codigo", "Es Modulo", "Modulo", "Modulo Orden", "ESTADO"]
    for col in needed_p:
        if col not in p.columns:
            p[col] = ""

    h["Catalogo"] = h["Catalogo"].map(norm)
    h["Codigo_Catalogo"] = h["Codigo_Catalogo"].map(norm)

    p["Catalogo"] = p["Catalogo"].map(norm)
    p["Codigo"] = p["Codigo"].map(norm)
    p["Es Modulo"] = p["Es Modulo"].map(norm)
    p["Modulo"] = p["Modulo"].map(norm)

    merged = h.merge(
        p[["Catalogo", "Codigo", "Es Modulo", "Modulo", "Modulo Orden"]],
        left_on=["Catalogo", "Codigo_Catalogo"],
        right_on=["Catalogo", "Codigo"],
        how="left",
    )

    merged["Es Modulo"] = merged["Es Modulo"].fillna("").map(norm)
    merged["Modulo"] = merged["Modulo"].fillna("").map(norm)

    merged.rename(
        columns={
            "codigo_referencia": "Codigo Referencia",
            "Codigo_Catalogo": "Codigo Catalogo",
            "PrestacionReferencia": "Prestacion Referencia",
            "PrestacionCatalogo": "Prestacion Catalogo",
            "Tipo Homologacion": "Tipo Homologacion",
            "ESTADO": "Estado",
            "fecha inicio vigencia": "Vigencia Inicio",
        },
        inplace=True,
    )

    return merged


def get_catalogos_for_convenio(catalogos_convenio: pd.DataFrame, convenio_name: str) -> list[str]:
    if catalogos_convenio.empty:
        return []

    df = catalogos_convenio.copy()
    if "Convenio" not in df.columns or "Catalogo" not in df.columns:
        return []

    mask = df["Convenio"].map(norm).str.upper() == convenio_name.upper()
    cats = sorted({norm(v) for v in df.loc[mask, "Catalogo"].tolist() if norm(v)})
    return cats


def get_modulo_drilldown(modulo_prestacion: pd.DataFrame, modulos_reglas: pd.DataFrame, modulo_name: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    mp = modulo_prestacion.copy()
    mr = modulos_reglas.copy()

    for col in ["Modulo", "CodigoPrestacionReferencia", "PrestacionReferencia", "TipoInclusion", "Tope", "ESTADO"]:
        if col not in mp.columns:
            mp[col] = ""

    for col in ["ReglaModulada", "TipoInclusion", "Orden", "Estado"]:
        if col not in mr.columns:
            mr[col] = ""

    mp_mask = mp["Modulo"].map(norm).str.upper() == modulo_name.upper()
    mp_out = mp.loc[mp_mask, ["Modulo", "CodigoPrestacionReferencia", "PrestacionReferencia", "TipoInclusion", "Tope", "ESTADO"]]

    modulo_col_mr = "Modulo" if "Modulo" in mr.columns else mr.columns[0]
    mr_mask = mr[modulo_col_mr].map(norm).str.upper() == modulo_name.upper()
    mr_out = mr.loc[mr_mask, [modulo_col_mr, "ReglaModulada", "TipoInclusion", "Orden", "Estado"]]
    mr_out.rename(columns={modulo_col_mr: "Modulo"}, inplace=True)

    return mp_out.reset_index(drop=True), mr_out.reset_index(drop=True)


def main() -> None:
    st.set_page_config(page_title="Baseline Convenios - San Jose", layout="wide")

    st.title("Baseline de Configuracion de Convenios")
    st.caption(f"Centro baseline: {BASELINE_CENTER}")

    excel_path = Path(st.text_input("Ruta Excel", value=DEFAULT_EXCEL)).expanduser()
    if not excel_path.exists():
        st.error("No se encontro el archivo Excel en la ruta indicada.")
        return

    with st.spinner("Cargando y normalizando hojas del template..."):
        data = load_data(excel_path)

    convenios = data["convenios"]
    convenios_planes = data["convenios_planes"]
    homologaciones = data["homologaciones"]
    prestaciones_catalogos = data["prestaciones_catalogos"]
    catalogos_convenio = data["catalogos_convenio"]
    modulo_prestacion = data["modulo_prestacion"]
    modulos_reglas = data["modulos_reglas"]

    if convenios.empty:
        st.error("No se pudieron leer datos de la hoja Convenios.")
        return

    # Normalization for filters
    if "tipo convenio" not in convenios.columns:
        convenios["tipo convenio"] = ""
    if "nombre" not in convenios.columns:
        convenios["nombre"] = ""

    convenios["tipo convenio"] = convenios["tipo convenio"].map(norm)
    convenios["nombre"] = convenios["nombre"].map(norm)

    st.subheader("Filtros superiores")
    c1, c2, c3, c4, c5 = st.columns([1.2, 1.2, 1.2, 1.4, 1.0])

    tipos = sorted([t for t in convenios["tipo convenio"].unique().tolist() if t])
    tipo_sel = c1.selectbox("Tipo de convenio", options=tipos, index=0 if tipos else None)

    conv_options = sorted(
        convenios.loc[convenios["tipo convenio"] == tipo_sel, "nombre"].dropna().map(norm).unique().tolist()
    )
    convenio_sel = c2.selectbox("Convenio", options=conv_options, index=0 if conv_options else None)

    cp = convenios_planes.copy()
    for col in ["Convenio", "Financiador", "Plan", "Estado"]:
        if col not in cp.columns:
            cp[col] = ""
        cp[col] = cp[col].map(norm)

    cp_conv = cp.loc[cp["Convenio"].str.upper() == norm(convenio_sel).upper()].copy()

    fin_options = sorted([v for v in cp_conv["Financiador"].unique().tolist() if v])
    fin_sel = c3.selectbox("Financiador asociado", options=fin_options, index=0 if fin_options else None)

    plan_options = sorted(
        [
            v
            for v in cp_conv.loc[cp_conv["Financiador"].str.upper() == norm(fin_sel).upper(), "Plan"].unique().tolist()
            if v
        ]
    )
    plan_sel = c4.selectbox("Plan", options=plan_options, index=0 if plan_options else None)

    only_modulo = c5.checkbox("Solo modulos", value=False)

    st.markdown("---")

    # Context cards
    cards = st.columns(4)
    cards[0].metric("Centro", BASELINE_CENTER)
    cards[1].metric("Tipo convenio", norm(tipo_sel) or "-")
    cards[2].metric("Convenio", norm(convenio_sel) or "-")
    cards[3].metric("Financiador / Plan", f"{norm(fin_sel) or '-'} / {norm(plan_sel) or '-'}")

    # Catalogs to filter homologaciones
    catalogos = get_catalogos_for_convenio(catalogos_convenio, norm(convenio_sel))
    if not catalogos:
        # fallback: if no mapping exists, assume convenio name as catalog name
        if norm(convenio_sel):
            catalogos = [norm(convenio_sel)]

    hbase = merge_homologaciones_with_modulo(homologaciones, prestaciones_catalogos)
    hbase["Catalogo"] = hbase["Catalogo"].map(norm)

    hview = hbase.loc[hbase["Catalogo"].isin(catalogos)].copy()

    # Extra filters on top for readability
    f1, f2, f3 = st.columns([1.0, 1.0, 2.2])
    estado_opts = sorted([v for v in hview["Estado"].map(norm).unique().tolist() if v])
    tipo_h_opts = sorted([v for v in hview["Tipo Homologacion"].map(norm).unique().tolist() if v])

    estado_sel = f1.multiselect("Estado homologacion", options=estado_opts, default=estado_opts)
    tipo_h_sel = f2.multiselect("Tipo homologacion", options=tipo_h_opts, default=tipo_h_opts)
    search_txt = f3.text_input("Buscar por codigo o nombre", value="")

    if estado_sel:
        hview = hview.loc[hview["Estado"].map(norm).isin(estado_sel)]
    if tipo_h_sel:
        hview = hview.loc[hview["Tipo Homologacion"].map(norm).isin(tipo_h_sel)]
    if only_modulo:
        hview = hview.loc[hview["Es Modulo"].map(norm).str.upper() == "SI"]

    search_txt_norm = norm(search_txt).upper()
    if search_txt_norm:
        mask = (
            hview["Codigo Referencia"].map(norm).str.upper().str.contains(search_txt_norm, na=False)
            | hview["Codigo Catalogo"].map(norm).str.upper().str.contains(search_txt_norm, na=False)
            | hview["Prestacion Referencia"].map(norm).str.upper().str.contains(search_txt_norm, na=False)
            | hview["Prestacion Catalogo"].map(norm).str.upper().str.contains(search_txt_norm, na=False)
            | hview["Modulo"].map(norm).str.upper().str.contains(search_txt_norm, na=False)
        )
        hview = hview.loc[mask]

    st.subheader("Homologaciones configuradas (baseline San Jose)")
    k1, k2, k3 = st.columns(3)
    k1.metric("Homologaciones visibles", len(hview))
    k2.metric("Catalogos vinculados", len(catalogos))
    k3.metric("Filas modulo (SI)", int((hview["Es Modulo"].map(norm).str.upper() == "SI").sum()))

    cols_show = [
        "Catalogo",
        "Codigo Referencia",
        "Prestacion Referencia",
        "Codigo Catalogo",
        "Prestacion Catalogo",
        "Tipo Homologacion",
        "Es Modulo",
        "Modulo",
        "Estado",
        "Vigencia Inicio",
    ]
    st.dataframe(hview[cols_show], use_container_width=True, height=420)

    st.subheader("Drilldown de modulos")
    modulos_visibles = sorted(
        [
            m
            for m in hview.loc[hview["Es Modulo"].map(norm).str.upper() == "SI", "Modulo"].map(norm).unique().tolist()
            if m
        ]
    )

    if not modulos_visibles:
        st.info("No hay prestaciones modulares en la seleccion actual.")
        return

    modulo_sel = st.selectbox("Seleccionar modulo", options=modulos_visibles, index=0)
    mp_out, mr_out = get_modulo_drilldown(modulo_prestacion, modulos_reglas, modulo_sel)

    d1, d2 = st.columns(2)
    with d1:
        st.markdown("Prestaciones dentro del modulo")
        st.dataframe(mp_out, use_container_width=True, height=320)
    with d2:
        st.markdown("Reglas moduladas del modulo")
        st.dataframe(mr_out, use_container_width=True, height=320)


if __name__ == "__main__":
    main()
