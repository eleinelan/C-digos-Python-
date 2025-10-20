from flask import Flask, render_template
import subprocess
import os

app = Flask(__name__)
CAMINHO_CODIGOS = os.path.join(os.path.dirname(__file__), "codigos")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/automacao_fsist_recebidas")
def automacao_fsist_recebidas():
    subprocess.Popen(["python", os.path.join(CAMINHO_CODIGOS, "automacao_fsist_recebidas.py")])
    return "Script automacao_fsist_recebidas.py iniciado!"

@app.route("/nfse_bot")
def nfse_bot():
    subprocess.Popen(["python", os.path.join(CAMINHO_CODIGOS, "nfse_bot.py")])
    return "Script nfse_bot.py iniciado!"

@app.route("/nfsenacional_emitidasrecebidas")
def nfsenacional_emitidasrecebidas():
    subprocess.Popen(["python", os.path.join(CAMINHO_CODIGOS, "nfsenacional_emitidasrecebidas.py")])
    return "Script nfsenacional_emitidasrecebidas.py iniciado!"

@app.route("/osasco_fluxo")
def osasco_fluxo():
    subprocess.Popen(["python", os.path.join(CAMINHO_CODIGOS, "osasco_fluxo.py")])
    return "Script osasco_fluxo.py iniciado!"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
