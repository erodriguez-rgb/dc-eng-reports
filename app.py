#!/usr/bin/env python3
"""
DC Eng Reports - Servidor Flask para generación de informes DOCX
"""

from flask import Flask, request, jsonify, send_file, Response
from flask_cors import CORS
import json, os, sys, io, tempfile, base64
from generate_report import generar_informe
from generate_monthly_report import generar_informe_mensual

app = Flask(__name__)
CORS(app)  # Habilitar CORS para peticiones desde Base44 (app.base44.com)

# Token de seguridad
API_TOKEN = os.environ.get('API_TOKEN', 'changeme')

def check_auth(req):
    token = req.headers.get('X-API-Token') or req.args.get('token')
    return token == API_TOKEN

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'dc-eng-reports', 'version': '2.0'})

@app.route('/generate', methods=['POST'])
def generate():
    """Genera informe DIARIO"""
    if not check_auth(request):
        return jsonify({'error': 'Unauthorized'}), 401

    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No data provided'}), 400

        output_path = generar_informe(data)

        with open(output_path, 'rb') as f:
            docx_bytes = f.read()

        docx_b64 = base64.b64encode(docx_bytes).decode('utf-8')
        filename = f"Informe_Diario_{data.get('fecha', 'fecha')}.docx".replace('/', '-')

        return jsonify({
            'success': True,
            'filename': filename,
            'docx_base64': docx_b64
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/generar-informe-mensual', methods=['POST'])
def generar_mensual():
    """Genera informe MENSUAL completo como DOCX binario"""
    if not check_auth(request):
        return jsonify({'error': 'Unauthorized'}), 401

    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No data provided'}), 400

        buf = generar_informe_mensual(data)

        numero = data.get('numero_informe', 1)
        nombre_proy = data.get('nombre_proyecto', 'Proyecto')[:30].replace(' ', '_').replace('/', '-')
        filename = f"Informe_Mensual_{numero}_{nombre_proy}.docx"

        return Response(
            buf.read(),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            headers={
                'Content-Disposition': f'attachment; filename="{filename}"',
                'X-Generated-By': 'DC Eng Reports v2'
            }
        )

    except Exception as e:
        import traceback
        print(f"[ERROR] generar-informe-mensual: {e}", flush=True)
        traceback.print_exc()
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
