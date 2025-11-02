from flask import Flask, jsonify, request, send_from_directory
import pandas as pd
import json
import os

app = Flask(__name__)
CAMINHO_PLANILHA = "demandas.xlsx"  # sua planilha Excel
STATUS_FILE = "status_coffees.json"

# Inicializa arquivo de status se não existir
if not os.path.exists(STATUS_FILE):
    with open(STATUS_FILE, "w") as f:
        json.dump({}, f)

def ler_status():
    with open(STATUS_FILE, "r") as f:
        return json.load(f)

def salvar_status(status):
    with open(STATUS_FILE, "w") as f:
        json.dump(status, f)

def parse_itens_field(valor):
    if pd.isna(valor):
        return []
    linhas = [l.strip() for l in str(valor).splitlines() if l.strip()]
    out = []
    for linha in linhas:
        parts = linha.split("|")
        if len(parts) >= 2:
            quantidade, item = parts[0].strip(), parts[1].strip()
        else:
            quantidade, item = "", linha
        out.append({"quantidade": quantidade, "item": item})
    return out

@app.route("/dados")
def dados():
    if not os.path.exists(CAMINHO_PLANILHA):
        return jsonify({"error": "Arquivo não encontrado"}), 404

    df = pd.read_excel(CAMINHO_PLANILHA, engine="openpyxl")
    status = ler_status()

    registros = []
    for idx, row in df.iterrows():
        r_id = str(row.get("ID", idx))  # ID único
        registros.append({
            "id": r_id,
            "data": str(row.get("Data","")),
            "horario": str(row.get("Hora","")),
            "local": str(row.get("Local","")),
            "secretaria": str(row.get("Secretaria","")),
            "diretor": str(row.get("Diretor","")),
            "pax": str(row.get("Pax","")),
            "observacoes": str(row.get("Observações","")),
            "itens": parse_itens_field(row.get("Itens","")),
            "status": str(row.get("Status","")),
            "retirado": status.get(r_id, False)
        })
    return jsonify(registros)

@app.route("/confirmar", methods=["POST"])
def confirmar():
    data = request.json
    r_id = data.get("id")
    if not r_id:
        return jsonify({"error":"ID não informado"}),400
    status = ler_status()
    status[r_id] = True
    salvar_status(status)
    return jsonify({"success": True})

@app.route("/")
def index():
    return send_from_directory(".", "index.html")

@app.route("/<path:filename>")
def static_files(filename):
    if os.path.exists(filename):
        return send_from_directory(".", filename)
    return "Arquivo não encontrado", 404

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
