#!/usr/bin/env python3
"""
DC Eng Reports - Servidor Flask para generación de informes DOCX
"""

from flask import Flask, request, jsonify, send_file
import json, os, sys, io, tempfile, base64
from generate_report import generar_informe

app = Flask(__name__)

# Token de seguridad para proteger el endpoint
API_TOKEN = os.environ.get('API_TOKEN', 'changeme')

def check_auth(req):
    token = req.headers.get('X-API-Token') or req.args.get('token')
    return token == API_TOKEN

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'dc-eng-reports'})

@app.route('/generate', methods=['POST'])
def generate():
    if not check_auth(request):
        return jsonify({'error': 'Unauthorized'}), 401

    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No data provided'}), 400

        output_path = generar_informe(data)

        # Leer el archivo generado y retornarlo como base64
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

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
