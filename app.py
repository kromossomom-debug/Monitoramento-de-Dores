import sys
import os
import subprocess

# ==============================================================================
# 1. VERIFICAÇÃO E AUTO-INSTALAÇÃO DE DEPENDÊNCIAS
# ==============================================================================
REQUIRED_PACKAGES = [
    ("flask", "Flask>=3.0.0"),
    ("openpyxl", "openpyxl>=3.1.2")
]

def ensure_dependencies_installed():
    """
    Verifica se todas as dependências estão instaladas.
    Se faltar alguma biblioteca, instala automaticamente via pip antes de prosseguir.
    """
    missing = []
    for module_name, package_spec in REQUIRED_PACKAGES:
        try:
            __import__(module_name)
        except ImportError:
            missing.append(package_spec)

    if missing:
        print("=" * 70)
        print(" [AUTO-SETUP] Verificando ambiente Python...")
        print(" Foram identificadas dependências ausentes necessárias para o sistema:")
        for pkg in missing:
            print(f"   -> {pkg}")
        print("\n Instalando componentes automaticamente via pip... Aguarde um instante.")
        print("=" * 70)

        req_path = os.path.join(os.path.dirname(__file__), "requirements.txt")
        try:
            if os.path.exists(req_path):
                cmd = [sys.executable, "-m", "pip", "install", "-r", req_path]
            else:
                cmd = [sys.executable, "-m", "pip", "install", *missing]
            subprocess.check_call(cmd)
            print("\n [AUTO-SETUP] Todas as bibliotecas foram instaladas com sucesso!\n")
        except Exception as err:
            print(f"\n [ERRO] Falha ao instalar dependências automaticamente: {err}")
            print(f" Execute manualmente no terminal: pip install -r requirements.txt")
            sys.exit(1)

# Executa verificação imediatamente antes de qualquer import de terceiros
ensure_dependencies_installed()

# ==============================================================================
# 2. IMPORTS PRINCIPAIS DA APLICAÇÃO
# ==============================================================================
import datetime
import webbrowser
import threading
from flask import Flask, render_template, request, jsonify, send_file
import excel_manager

app = Flask(__name__)
app.config["JSON_AS_ASCII"] = False  # Suporte total a UTF-8

# ==============================================================================
# 3. ROTAS DA API & INTERFACE WEB
# ==============================================================================
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/status", methods=["GET"])
def get_status():
    file_path = excel_manager.get_active_excel_path()
    exists = os.path.exists(file_path)
    file_size = os.path.getsize(file_path) if exists else 0
    mtime = datetime.datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%d/%m/%Y %H:%M:%S") if exists else None
    
    backups = []
    if os.path.exists(excel_manager.BACKUP_DIR):
        backups = [f for f in os.listdir(excel_manager.BACKUP_DIR) if f.endswith(".xlsx")]
    
    return jsonify({
        "file_path": file_path,
        "exists": exists,
        "file_size_kb": round(file_size / 1024, 1),
        "last_modified": mtime,
        "backup_count": len(backups)
    })

@app.route("/api/parameters", methods=["GET"])
def get_parameters():
    try:
        params = excel_manager.get_parameters()
        return jsonify({"success": True, "data": params})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/records", methods=["GET"])
def get_records():
    try:
        records = excel_manager.get_all_records()
        
        status_filter = request.args.get("status")
        prioridade_filter = request.args.get("prioridade")
        setor_filter = request.args.get("setor")
        categoria_filter = request.args.get("categoria")
        filial_filter = request.args.get("filial")
        situacao_sla_filter = request.args.get("situacao_sla")
        search_query = request.args.get("search", "").strip().lower()

        filtered = []
        for r in records:
            if status_filter and r.get("status") != status_filter:
                continue
            if prioridade_filter and r.get("prioridade") != prioridade_filter:
                continue
            if setor_filter and (r.get("setor_responsavel") != setor_filter and r.get("setor_impactado") != setor_filter):
                continue
            if categoria_filter and r.get("categoria_dor") != categoria_filter:
                continue
            if filial_filter and r.get("filial") != filial_filter:
                continue
            if situacao_sla_filter and r.get("situacao_sla") != situacao_sla_filter:
                continue
            
            if search_query:
                haystack = " ".join([
                    str(r.get("id") or ""),
                    str(r.get("responsavel") or ""),
                    str(r.get("setor_responsavel") or ""),
                    str(r.get("setor_impactado") or ""),
                    str(r.get("filial") or ""),
                    str(r.get("categoria_dor") or ""),
                    str(r.get("descricao_problema") or ""),
                    str(r.get("causa_raiz") or ""),
                    str(r.get("impacto_negocio") or ""),
                    str(r.get("plano_acao") or ""),
                    str(r.get("nota_fiscal") or ""),
                    str(r.get("numero_chamado") or ""),
                    str(r.get("responsavel_solucao") or ""),
                ]).lower()
                if search_query not in haystack:
                    continue

            filtered.append(r)

        return jsonify({"success": True, "count": len(filtered), "data": filtered})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/records/<int:record_id>", methods=["GET"])
def get_record(record_id):
    try:
        records = excel_manager.get_all_records()
        for r in records:
            if r["id"] == record_id:
                return jsonify({"success": True, "data": r})
        return jsonify({"success": False, "error": f"Registro #{record_id} não encontrado."}), 404
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/kpis", methods=["GET"])
def get_kpis():
    try:
        records = excel_manager.get_all_records()
        kpis = excel_manager.calculate_kpis(records)
        return jsonify({"success": True, "data": kpis})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/records", methods=["POST"])
def create_record():
    try:
        payload = request.get_json(force=True)
        if not payload:
            return jsonify({"success": False, "message": "Nenhum dado recebido."}), 400
        
        if not payload.get("descricao_problema") or not str(payload.get("descricao_problema")).strip():
            return jsonify({"success": False, "message": "O campo 'Descrição do Problema' é obrigatório."}), 400
        
        res = excel_manager.add_record(payload)
        status_code = 200 if res.get("success") else 400
        return jsonify(res), status_code
    except Exception as e:
        return jsonify({"success": False, "message": f"Erro interno ao cadastrar: {str(e)}"}), 500

@app.route("/api/records/<int:record_id>", methods=["PUT"])
def update_record(record_id):
    try:
        payload = request.get_json(force=True)
        if not payload:
            return jsonify({"success": False, "message": "Nenhum dado recebido."}), 400
        
        res = excel_manager.update_record(record_id, payload)
        status_code = 200 if res.get("success") else 400
        return jsonify(res), status_code
    except Exception as e:
        return jsonify({"success": False, "message": f"Erro interno ao atualizar: {str(e)}"}), 500

@app.route("/api/download-excel", methods=["GET"])
def download_excel():
    try:
        file_path = excel_manager.get_active_excel_path()
        return send_file(
            file_path,
            as_attachment=True,
            download_name="Monitoramento_Dores_Atualizado.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/open-excel", methods=["POST"])
def open_excel():
    try:
        file_path = excel_manager.get_active_excel_path()
        if os.path.exists(file_path):
            if sys.platform.startswith("win"):
                os.startfile(file_path)
            elif sys.platform.startswith("darwin"):
                subprocess.call(["open", file_path])
            else:
                subprocess.call(["xdg-open", file_path])
            return jsonify({"success": True, "message": f"Abrindo {os.path.basename(file_path)} no Excel."})
        else:
            return jsonify({"success": False, "message": "Arquivo Excel não encontrado."}), 404
    except Exception as e:
        return jsonify({"success": False, "message": f"Não foi possível abrir o Excel: {str(e)}"}), 500

def open_browser():
    try:
        webbrowser.open_new("http://127.0.0.1:5000")
    except Exception:
        pass

if __name__ == "__main__":
    active_path = excel_manager.get_active_excel_path()
    print("=" * 70)
    print("   SISTEMA DE MAPEAMENTO DE DORES - PAINEL EXECUTIVO & OPERACIONAL")
    print(f"   Base de Dados Excel ativa: {active_path}")
    print("   Servidor iniciado em: http://127.0.0.1:5000")
    print("   Pressione CTRL+C no terminal para encerrar.")
    print("=" * 70)

    # Abre o navegador automaticamente em segundo plano
    threading.Timer(1.2, open_browser).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
