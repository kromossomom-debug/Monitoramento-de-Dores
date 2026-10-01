import sys
import os
import subprocess
import io
import csv

# ==============================================================================
# 1. VERIFICAÇÃO E AUTO-INSTALAÇÃO DE DEPENDÊNCIAS
# ==============================================================================
REQUIRED_PACKAGES = [
    ("flask", "Flask>=3.0.0"),
    ("openpyxl", "openpyxl>=3.1.2"),
    ("reportlab", "reportlab>=4.0.0")
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

ensure_dependencies_installed()

# ==============================================================================
# 2. IMPORTS PRINCIPAIS DA APLICAÇÃO
# ==============================================================================
import datetime
import webbrowser
import threading
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory, Response
import excel_manager
import pdf_generator

app = Flask(__name__)
app.config["JSON_AS_ASCII"] = False  # Suporte total a UTF-8
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # Suporte a uploads de até 50MB

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

    # Count total attachments stored in uploads/
    total_attachments_stored = 0
    if os.path.exists(excel_manager.UPLOADS_DIR):
        for root, dirs, files in os.walk(excel_manager.UPLOADS_DIR):
            total_attachments_stored += len(files)

    return jsonify({
        "file_path": file_path,
        "exists": exists,
        "file_size_kb": round(file_size / 1024, 1),
        "last_modified": mtime,
        "backup_count": len(backups),
        "total_attachments": total_attachments_stored
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
        has_anexos_filter = request.args.get("has_anexos")
        search_query = request.args.get("search", "").strip().lower()

        # Date range filtering
        data_inicio = request.args.get("data_inicio")
        data_fim = request.args.get("data_fim")

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
            if has_anexos_filter == "1" and r.get("total_anexos", 0) == 0:
                continue

            # Date Range check
            reg_iso = r.get("data_registro_iso")
            if reg_iso:
                if data_inicio and reg_iso < data_inicio:
                    continue
                if data_fim and reg_iso > data_fim:
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
                    str(r.get("anexos_nomes") or "")
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

        # Optional date filtering for KPIs
        data_inicio = request.args.get("data_inicio")
        data_fim = request.args.get("data_fim")

        if data_inicio or data_fim:
            filtered_for_kpi = []
            for r in records:
                reg_iso = r.get("data_registro_iso")
                if reg_iso:
                    if data_inicio and reg_iso < data_inicio:
                        continue
                    if data_fim and reg_iso > data_fim:
                        continue
                filtered_for_kpi.append(r)
            records = filtered_for_kpi

        kpis = excel_manager.calculate_kpis(records)
        return jsonify({"success": True, "data": kpis})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/records", methods=["POST"])
def create_record():
    """
    Creates a new demand. Accepts either JSON or multipart/form-data with attached files.
    """
    try:
        files = []
        if request.content_type and "multipart/form-data" in request.content_type:
            payload = request.form.to_dict()
            # Collect uploaded files from 'anexos' or 'files'
            files = request.files.getlist("anexos") or request.files.getlist("files")
        else:
            payload = request.get_json(force=True) or {}

        if not payload:
            return jsonify({"success": False, "message": "Nenhum dado recebido."}), 400

        if not payload.get("descricao_problema") or not str(payload.get("descricao_problema")).strip():
            return jsonify({"success": False, "message": "O campo 'Descrição do Problema' é obrigatório."}), 400

        res = excel_manager.add_record(payload, files=files)
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

@app.route("/api/records/<int:record_id>/quick-complete", methods=["POST"])
def quick_complete(record_id):
    """Marks a demand as concluded with 1 click."""
    try:
        obs = request.json.get("observacoes", "") if request.is_json else ""
        res = excel_manager.quick_complete_record(record_id, obs)
        status_code = 200 if res.get("success") else 400
        return jsonify(res), status_code
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

# ==============================================================================
# 4. GESTÃO DE ANEXOS (UPLOAD, LISTAGEM, DOWNLOAD, EXCLUSÃO)
# ==============================================================================
@app.route("/api/records/<int:record_id>/attachments", methods=["GET"])
def get_attachments(record_id):
    """Lists attachments for a specific record."""
    try:
        attachments = excel_manager.get_record_attachments(record_id)
        return jsonify({"success": True, "count": len(attachments), "data": attachments})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/records/<int:record_id>/attachments", methods=["POST"])
def upload_attachments(record_id):
    """Uploads one or more files to an existing record."""
    try:
        files = request.files.getlist("anexos") or request.files.getlist("files")
        if not files:
            return jsonify({"success": False, "message": "Nenhum arquivo enviado."}), 400

        saved = excel_manager.save_record_attachments(record_id, files)
        return jsonify({
            "success": True,
            "message": f"{len(files)} arquivo(s) anexado(s) com sucesso!",
            "data": saved
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"Erro ao anexar arquivo: {str(e)}"}), 500

@app.route("/api/attachments/<int:record_id>/<path:filename>", methods=["GET"])
def serve_attachment(record_id, filename):
    """Serves an attached file. If download=1 query is passed, forces download. Otherwise previews inline."""
    try:
        rec_dir = os.path.join(excel_manager.UPLOADS_DIR, str(record_id))
        safe_name = os.path.basename(filename)
        file_path = os.path.join(rec_dir, safe_name)
        if not os.path.exists(file_path):
            return "Arquivo não encontrado.", 404

        as_download = request.args.get("download") == "1"
        return send_from_directory(
            rec_dir,
            safe_name,
            as_attachment=as_download,
            download_name=safe_name
        )
    except Exception as e:
        return f"Erro ao acessar arquivo: {str(e)}", 500

@app.route("/api/records/<int:record_id>/attachments/<path:filename>", methods=["DELETE"])
def delete_attachment(record_id, filename):
    """Deletes an attached file from a record."""
    try:
        success = excel_manager.delete_record_attachment(record_id, filename)
        if success:
            return jsonify({"success": True, "message": "Anexo removido com sucesso."})
        return jsonify({"success": False, "message": "Arquivo não encontrado."}), 404
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

# ==============================================================================
# 5. EXPORTAÇÕES E INTEGRAÇÃO NATIVA
# ==============================================================================
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

@app.route("/api/export-csv", methods=["GET"])
def export_csv():
    """Exports records as a formatted CSV with UTF-8 BOM for Microsoft Excel."""
    try:
        records = excel_manager.get_all_records()

        # Query filters
        status_filter = request.args.get("status")
        prioridade_filter = request.args.get("prioridade")
        filial_filter = request.args.get("filial")

        filtered = []
        for r in records:
            if status_filter and r.get("status") != status_filter:
                continue
            if prioridade_filter and r.get("prioridade") != prioridade_filter:
                continue
            if filial_filter and r.get("filial") != filial_filter:
                continue
            filtered.append(r)

        output = io.StringIO()
        # UTF-8 BOM so Excel opens with proper accents
        output.write('\ufeff')
        writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_MINIMAL)

        writer.writerow([
            "ID", "Data Registro", "Data Ocorrência", "Responsável", "Setor Responsável",
            "Setor Impactado", "Filial", "Categoria da Dor", "Nota Fiscal", "Série",
            "Descrição do Problema", "Causa Raiz", "Impacto no Negócio", "Qtd. Da Nota (kg)", "Valor Notas (R$)",
            "Prioridade", "Status", "SLA (dias)", "Data Limite SLA", "Situação SLA",
            "Dias em Aberto", "Plano de Ação", "Prazo", "Data Conclusão", "Responsável Solução",
            "Nº Chamado", "Recorrente?", "Total Anexos", "Nomes dos Anexos"
        ])

        for r in filtered:
            writer.writerow([
                r.get("id"),
                r.get("data_registro_br", ""),
                r.get("data_ocorrencia_br", ""),
                r.get("responsavel", ""),
                r.get("setor_responsavel", ""),
                r.get("setor_impactado", ""),
                r.get("filial", ""),
                r.get("categoria_dor", ""),
                r.get("nota_fiscal", ""),
                r.get("serie", ""),
                r.get("descricao_problema", ""),
                r.get("causa_raiz", ""),
                r.get("impacto_negocio", ""),
                r.get("qtd_nota_kg_mask", ""),
                r.get("valor_notas_contabil", "0,00"),
                r.get("prioridade", ""),
                r.get("status", ""),
                r.get("sla_dias", ""),
                r.get("data_limite_sla_br", ""),
                r.get("situacao_sla", ""),
                r.get("dias_aberto", ""),
                r.get("plano_acao", ""),
                r.get("prazo_br", ""),
                r.get("data_conclusao_br", ""),
                r.get("responsavel_solucao", ""),
                r.get("numero_chamado", ""),
                r.get("recorrente", ""),
                r.get("total_anexos", 0),
                r.get("anexos_nomes", "")
            ])

        csv_data = output.getvalue().encode('utf-8-sig')
        return Response(
            csv_data,
            mimetype="text/csv",
            headers={"Content-disposition": "attachment; filename=relatorio_dores_filtrado.csv"}
        )
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/relatorio-pdf", methods=["GET"])
def gerar_relatorio_pdf():
    """Generates an executive diagnostic PDF report filtered by date range and criteria."""
    try:
        records = excel_manager.get_all_records()

        data_inicio = request.args.get("data_inicio")
        data_fim = request.args.get("data_fim")
        status_filter = request.args.get("status")
        filial_filter = request.args.get("filial")
        periodo_label = request.args.get("periodo_label")
        as_view = request.args.get("view") == "1"

        filtered = []
        for r in records:
            if status_filter and r.get("status") != status_filter:
                continue
            if filial_filter and r.get("filial") != filial_filter:
                continue

            reg_iso = r.get("data_registro_iso")
            if reg_iso:
                if data_inicio and reg_iso < data_inicio:
                    continue
                if data_fim and reg_iso > data_fim:
                    continue
            filtered.append(r)

        pdf_bytes = pdf_generator.build_pdf_report(
            records=filtered,
            data_inicio=data_inicio,
            data_fim=data_fim,
            periodo_label=periodo_label
        )

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        download_name = f"Relatorio_Executivo_Dores_{ts}.pdf"

        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=(not as_view),
            download_name=download_name
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
    print(f"   Diretório de Anexos: {excel_manager.UPLOADS_DIR}")
    print("   Servidor iniciado em: http://127.0.0.1:5000")
    print("   Pressione CTRL+C no terminal para encerrar.")
    print("=" * 70)

    threading.Timer(1.2, open_browser).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
