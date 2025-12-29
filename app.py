import streamlit as st
import subprocess
import sys
import os
from pathlib import Path
from datetime import datetime

# ---------------------------
# Configuración de la app
# ---------------------------
st.set_page_config(
    page_title="Descarga MAEMP",
    page_icon="📥",
    layout="centered"
)

st.title("📥 Descarga INFORME MAEMP")
st.write(
    "Ejecuta manualmente el proceso de descarga. "
    "Esta versión es **100% compatible con Streamlit Cloud**."
)

# ---------------------------
# Paths
# ---------------------------
BASE_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = BASE_DIR / "PersonaActivo.py"

# ---------------------------
# Estado inicial
# ---------------------------
if "history" not in st.session_state:
    st.session_state.history = []

# ---------------------------
# Lógica BACKEND (NO Streamlit aquí)
# ---------------------------
def run_download_capture():
    """
    Ejecuta el script externo y captura stdout/stderr.
    Devuelve un dict con resultado, logs y metadata.
    """
    start_time = datetime.now()
    result_path = ""
    output = ""

    if not SCRIPT_PATH.exists():
        return {
            "time": start_time,
            "success": False,
            "result": "",
            "log": f"ERROR: Script no encontrado: {SCRIPT_PATH}"
        }

    cmd = [sys.executable, str(SCRIPT_PATH)]

    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=600
        )
        output = (proc.stdout or "") + (proc.stderr or "")
    except subprocess.TimeoutExpired:
        return {
            "time": start_time,
            "success": False,
            "result": "",
            "log": "ERROR: Timeout expirado durante la ejecución."
        }

    # Buscar RESULT_PATH::
    for line in output.splitlines()[::-1]:
        if line.strip().startswith("RESULT_PATH::"):
            result_path = line.split("RESULT_PATH::", 1)[1].strip()
            break

    return {
        "time": start_time,
        "success": bool(result_path),
        "result": result_path,
        "log": output
    }

# ---------------------------
# UI – Acción principal
# ---------------------------
if st.button("▶ Ejecutar descarga ahora"):
    with st.spinner("Ejecutando proceso..."):
        entry = run_download_capture()
        st.session_state.history.insert(0, entry)

    if entry["success"]:
        st.toast(f"Descarga completada: {os.path.basename(entry['result'])}")
    else:
        st.warning("No se detectó archivo descargado. Revisa los logs.")

# ---------------------------
# Última ejecución
# ---------------------------
st.markdown("---")
st.subheader("🕒 Última ejecución")

if st.session_state.history:
    last = st.session_state.history[0]

    st.write(f"**Hora:** {last['time'].strftime('%Y-%m-%d %H:%M:%S')}")

    if last["success"]:
        if os.path.exists(last["result"]):
            fname = os.path.basename(last["result"])
            st.success(f"Archivo generado: {fname}")
            st.write(f"Ruta: `{last['result']}`")
        else:
            st.warning("El archivo indicado no existe en el entorno.")
    else:
        st.error("La ejecución no generó archivo.")

    st.text_area(
        "Logs de la ejecución",
        last["log"],
        height=300
    )
else:
    st.info("Aún no se ha ejecutado el proceso.")

# ---------------------------
# Historial
# ---------------------------
st.markdown("---")
st.subheader("📜 Historial de ejecuciones")

if st.session_state.history:
    for h in st.session_state.history[:10]:
        status = "✅" if h["success"] else "❌"
        name = os.path.basename(h["result"]) if h["result"] else "sin archivo"
        st.write(
            f"{status} {h['time'].strftime('%Y-%m-%d %H:%M:%S')} — {name}"
        )
else:
    st.write("Sin historial aún.")

# ---------------------------
# Información técnica
# ---------------------------
st.markdown("---")
st.subheader("ℹ️ Notas técnicas")

st.write(
    "- Esta app **no usa schedulers ni procesos en background**.\n"
    "- Compatible con **Streamlit Cloud**.\n"
    "- Para ejecuciones automáticas, usa **GitHub Actions / Cron / Cloud Functions**.\n"
    "- Streamlit se usa solo como **interfaz de control y visualización**."
)
