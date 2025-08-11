# app.py
import streamlit as st
import subprocess, threading, sys, time, os
from pathlib import Path

st.set_page_config(page_title="Painel de Automação", layout="centered")
st.title("Automação — Atalhos rápidos")

# Scripts locais (mesma pasta do app)
ROOT = Path(__file__).parent
SCRIPTS = {
    "Automatizar Pedido": ROOT / "automatiza_pedido.py",
    "Gerar OF (RJ)":      ROOT / "automatiza_OF.py",
    "Processo Completo":  ROOT / "teste_completo.py",
}

# Estado
if "procs" not in st.session_state: st.session_state.procs = {}
if "logs" not in st.session_state:  st.session_state.logs  = {}
LOG_MAX = 12000

def run_script(name: str, script_path: Path, args=None):
    args = args or []
    if not script_path.exists():
        st.error(f"Arquivo não encontrado: {script_path}")
        return
    if name in st.session_state.procs and st.session_state.procs[name].poll() is None:
        st.warning(f"{name} já está em execução.")
        return

    st.session_state.logs[name] = ""
    cmd = [sys.executable, "-u", str(script_path), *args]
    creation = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        cwd=str(script_path.parent),
        creationflags=creation
    )
    st.session_state.procs[name] = proc

    def reader():
        for line in proc.stdout:
            st.session_state.logs[name] = (st.session_state.logs[name] + line)[-LOG_MAX:]
            time.sleep(0.02)
        code = proc.wait()
        st.session_state.logs[name] += f"\n[FINALIZADO] Código de saída: {code}\n"

    threading.Thread(target=reader, daemon=True).start()

def stop_script(name: str):
    proc = st.session_state.procs.get(name)
    if proc and proc.poll() is None:
        try:
            proc.terminate()
            time.sleep(0.8)
            if proc.poll() is None:
                proc.kill()
        except Exception as e:
            st.error(f"Erro ao parar {name}: {e}")

cols = st.columns(3)
labels = list(SCRIPTS.keys())
for i, label in enumerate(labels):
    with cols[i]:
        if st.button(label, use_container_width=True):
            run_script(label, SCRIPTS[label])

st.divider()
st.caption("Dica: deixe o Chrome/Siecon visível ao rodar scripts com PyAutoGUI.")

st.subheader("Tarefas")
if st.button("Atualizar status"):
    st.rerun()

for name in labels:
    running = name in st.session_state.procs and st.session_state.procs[name].poll() is None
    colA, colB = st.columns([3,1])
    colA.markdown(f"**{name}** — {'🟢 Em execução' if running else '⚪ Parado'}")
    if running and colB.button("Parar", key=f"stop_{name}"):
        stop_script(name)
    st.text_area("Logs", st.session_state.logs.get(name, ""), height=260, key=f"log_{name}")
    st.write("---")