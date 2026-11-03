"""Versión web del Auditor de Visitas (Streamlit). Usa el mismo motor que la aplicación de escritorio."""
import io
import os
import tempfile

import streamlit as st

import procesar_visitas as core

st.set_page_config(page_title="Auditor de Visitas · Ploom 2026", page_icon="📊", layout="centered")


def acceso_permitido():
    """Si hay una contraseña configurada en los 'secrets', se pide antes de mostrar la aplicación."""
    try:
        clave = st.secrets["PASSWORD"]
    except Exception:
        return True
    if st.session_state.get("acceso_ok"):
        return True
    txt = st.text_input("Contraseña", type="password")
    if txt:
        if txt == clave:
            st.session_state["acceso_ok"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta")
    return False


@st.cache_data(show_spinner=False)
def leer_base(contenido):
    return core.cargar_base_estancos(io.BytesIO(contenido))


def procesar(pdfs, base, formulas):
    lista, avisos_tabla, ilegibles, con_direccion = [], [], 0, 0
    barra = st.progress(0.0, text="Preparando…")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = os.path.join(tmp, "visita.pdf")
        for i, f in enumerate(pdfs, 1):
            barra.progress(i / len(pdfs), text=f"{i}/{len(pdfs)} · {f.name}")
            with open(ruta, "wb") as fh:
                fh.write(f.getbuffer())
            try:
                registro, avisos, dir_ok = core.procesar_pdf(ruta, base)
            except Exception as e:
                registro, avisos, dir_ok = None, [f"error al leer: {e}"], False
            if registro is None:
                ilegibles += 1
                avisos_tabla.append({"PDF": f.name, "Aviso": "No se pudo leer"})
                continue
            lista.append(registro)
            con_direccion += 1 if dir_ok else 0
            if avisos:
                avisos_tabla.append({"PDF": f.name, "Aviso": ", ".join(avisos)})
        if not lista:
            barra.empty()
            return None
        barra.progress(1.0, text="Generando el Excel…")
        salida = os.path.join(tmp, "Auditoria_Visitas_Total.xlsx")
        core.generar_excel(lista, salida, formulas)
        with open(salida, "rb") as fh:
            excel = fh.read()
    barra.empty()
    return {"excel": excel, "visitas": len(lista), "ilegibles": ilegibles, "avisos": avisos_tabla,
            "direcciones": con_direccion if base is not None else None}


if acceso_permitido():
    st.title("📊 Auditor de Visitas")
    st.caption("Campaña Ploom 2026 · Sube los informes y descarga el Excel con el dashboard.")

    pdfs = st.file_uploader("1 · Informes de visita (PDF)", type="pdf", accept_multiple_files=True)
    bbdd = st.file_uploader("2 · Base de datos de estancos (Excel, opcional)", type=["xlsx", "xlsm"])
    formulas = st.checkbox("Excel con fórmulas (se recalcula y el Resumen sigue los filtros)", value=True)
    st.caption("Consejo: sube los informes por tandas (por ejemplo, una semana cada vez). "
               "Los archivos no se guardan de forma permanente.")

    if st.button("Procesar y generar Excel", type="primary", disabled=not pdfs):
        base = None
        if bbdd is not None:
            try:
                base = leer_base(bbdd.getvalue())
                st.success(f"Base de estancos cargada: {base['n']} estancos.")
            except Exception as e:
                st.warning(f"No se pudo leer la base de datos ({e}). Se continúa sin direcciones.")
        with st.spinner("Procesando…"):
            st.session_state["resultado"] = procesar(pdfs, base, formulas)
        if st.session_state["resultado"] is None:
            st.error("No se pudo extraer datos de ningún PDF.")

    res = st.session_state.get("resultado")
    if res:
        c1, c2, c3 = st.columns(3)
        c1.metric("Visitas procesadas", res["visitas"])
        c2.metric("PDFs con avisos", len(res["avisos"]))
        c3.metric("Ilegibles", res["ilegibles"])
        if res["direcciones"] is not None:
            st.info(f"Direcciones encontradas: {res['direcciones']} de {res['visitas']} visitas.")
        st.download_button("⬇ Descargar Excel", data=res["excel"], file_name="Auditoria_Visitas_Total.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")
        if res["avisos"]:
            st.subheader("Avisos")
            st.dataframe(res["avisos"], use_container_width=True, hide_index=True)
