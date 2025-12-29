import streamlit as st
import subprocess
import sys
import os
from pathlib import Path
from datetime import datetime

# Scheduler y notificaciones
try:
    from apscheduler.schedulers.background import BackgroundScheduler
    APS_AVAILABLE = True
except Exception:
    APS_AVAILABLE = False

try:
    from plyer import notification
    PLYER_AVAILABLE = True
except Exception:
    PLYER_AVAILABLE = False

st.set_page_config(page_title="Descarga MAEMP", page_icon="📥")
st.title("📥 Ejecutar descarga INFORME MAEMP (headless)")
st.write("Ejecuta manualmente o programa ejecuciones periódicas. Ver logs y notificaciones aquí.")

# Paths
BASE_DIR = Path(__file__).resolve().parent
script_path = BASE_DIR / "PersonaActivo.py"

# Inicializar estado
if "last_result" not in st.session_state:
    st.session_state.last_result = ""
if "last_log" not in st.session_state:
    st.session_state.last_log = ""
if "history" not in st.session_state:
    st.session_state.history = []
if "scheduler_started" not in st.session_state:
    st.session_state.scheduler_started = False
if "scheduler_enabled" not in st.session_state:
    st.session_state.scheduler_enabled = False
if "scheduler" not in st.session_state:
    st.session_state.scheduler = None

def notify(title, msg):
    # Notificación en UI (persistente) + opcional sistema
    st.toast(msg) if hasattr(st, "toast") else st.info(msg)
    if PLYER_AVAILABLE:
        try:
            notification.notify(title=title, message=msg)
        except:
            pass

def run_download_capture():
    """Ejecuta el script y captura salida; devuelve (result_path, stdout)"""
    if not script_path.exists():
        return "", f"ERROR: Script no encontrado: {script_path}\n"
    cmd = [sys.executable, str(script_path)]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        output = proc.stdout + proc.stderr
    except subprocess.TimeoutExpired:
        output = "ERROR: Timeout expirado durante la ejecución del script.\n"
        return "", output
    # Buscar RESULT_PATH::
    result_path = ""
    for line in output.splitlines()[::-1]:
        if line.strip().startswith("RESULT_PATH::"):
            result_path = line.strip().split("RESULT_PATH::", 1)[1].strip()
            break
    return result_path, output

def run_and_update(notify_user=True):
    result_path, output = run_download_capture()
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = {"time": fecha, "result": result_path, "log": output}
    st.session_state.history.insert(0, entry)
    st.session_state.last_result = result_path
    st.session_state.last_log = output
    if result_path:
        msg = f"Descarga completada: {os.path.basename(result_path)}"
        if notify_user:
            notify("Descarga MAEMP", msg)
    else:
        if notify_user:
            notify("Descarga MAEMP", "No se detectó archivo descargado. Revisa logs.")
    return entry

# UI: controls
col1, col2 = st.columns([2,1])
with col1:
    if st.button("Ejecutar ahora"):
        with st.spinner("Ejecutando..."):
            entry = run_and_update(notify_user=True)
            st.text_area("Salida (última ejecución)", entry["log"], height=300)
with col2:
    st.write("Programación")
    if not APS_AVAILABLE:
        st.warning("APScheduler no instalado: instalar `apscheduler` para scheduler.")
    interval_minutes = st.number_input("Intervalo (minutos)", min_value=1, value=30, step=1)
    enable = st.checkbox("Habilitar ejecución periódica", value=st.session_state.scheduler_enabled)
    # actualizar estado
    st.session_state.scheduler_enabled = enable

# Scheduler management (persist across reruns)
if enable and APS_AVAILABLE:
    if not st.session_state.scheduler_started:
        scheduler = BackgroundScheduler()
        # job id único
        scheduler.add_job(lambda: run_and_update(notify_user=True), "interval", minutes=int(interval_minutes), id="maemp_job", replace_existing=True)
        scheduler.start()
        st.session_state.scheduler = scheduler
        st.session_state.scheduler_started = True
        st.success(f"Scheduler iniciado: cada {interval_minutes} minutos")
    else:
        # si ya estaba iniciado, actualizar el intervalo si cambió
        sched = st.session_state.scheduler
        try:
            job = sched.get_job("maemp_job")
            if job:
                job.reschedule("interval", minutes=int(interval_minutes))
        except Exception:
            pass
elif not enable:
    if st.session_state.scheduler_started and st.session_state.scheduler:
        try:
            st.session_state.scheduler.remove_all_jobs()
            st.session_state.scheduler.shutdown(wait=False)
        except Exception:
            pass
        st.session_state.scheduler = None
        st.session_state.scheduler_started = False
        st.success("Scheduler detenido")

# Mostrar último resultado y logs
st.markdown("---")
st.subheader("Última ejecución")
if st.session_state.history:
    latest = st.session_state.history[0]
    st.write(f"- Hora: {latest['time']}")
    if latest["result"]:
        if os.path.exists(latest["result"]):
            fname = os.path.basename(latest["result"])
            st.markdown(f"**Archivo:** [{fname}](file://{latest['result']})")
            st.write(f"Ruta: {latest['result']}")
        else:
            st.write("Archivo indicado no existe en disco.")
    else:
        st.write("No se encontró archivo descargado en la última ejecución.")
    st.text_area("Logs", latest["log"], height=300)
else:
    st.write("Aún no se han realizado ejecuciones.")

# Historial
st.markdown("---")
st.subheader("Historial (últimas ejecuciones)")
for h in st.session_state.history[:10]:
    status = "✅" if h["result"] else "❌"
    st.write(f"{status} {h['time']} — {os.path.basename(h['result']) if h['result'] else 'sin archivo'}")

# Instrucciones dependencias
st.markdown("---")
st.subheader("Requerimientos")
st.write("- `streamlit` (ya instalado)")
st.write("- Para scheduler: `apscheduler`")
st.write("- Notificaciones del sistema (opcional): `plyer`")
st.code("python -m pip install apscheduler plyer", language="bash")