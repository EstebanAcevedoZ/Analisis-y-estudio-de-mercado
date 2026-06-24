# ==========================================
# Archivo: app.py
# Descripción: Backend de Flask para conectar la UI con funciones_em.py
# ==========================================

from flask import Flask, render_template, request, send_file, jsonify
import os
import time
import threading
import webbrowser
# Import your core functions
from funciones_em import cargar_bases, crear_word

app = Flask(__name__)

# --- SMART CACHING SYSTEM ---
# This prevents reloading the 200MB Excel files every time you click "Generate".
# It only reloads if you change the file names/paths in the UI.
CACHE = {
    'paths': {
        'archivo_csv': None,
        'db_gi': None,
        'db_cu': None,
        'db_gi_comercial': None
    },
    'df_catastral': None,
    'dict_dfs': None
}

@app.route('/')
def index():
    # Renders the UI (templates/index.html)
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_report():
    data = request.json
    
    # 1. Extract DB Paths
    archivo_csv = data.get('archivo_datos')
    db_gi = data.get('db_gi')
    db_cu = data.get('db_cu')
    db_gi_comercial = data.get('db_gi_comercial')

    # 2. Check if we need to load/reload databases
    paths_changed = (
        CACHE['paths']['archivo_csv'] != archivo_csv or
        CACHE['paths']['db_gi'] != db_gi or
        CACHE['paths']['db_cu'] != db_cu or
        CACHE['paths']['db_gi_comercial'] != db_gi_comercial
    )

    if paths_changed or CACHE['dict_dfs'] is None:
        print("\n[INFO] Cargando bases de datos en memoria (Esto tomará unos segundos)...")
        try:
            df_catastral, dict_dfs = cargar_bases(
                archivo_csv=archivo_csv,
                db_gi=db_gi,
                db_cu=db_cu,
                db_gi_comercial=db_gi_comercial
            )
            # Update Cache
            CACHE['df_catastral'] = df_catastral
            CACHE['dict_dfs'] = dict_dfs
            CACHE['paths'] = {
                'archivo_csv': archivo_csv, 'db_gi': db_gi, 
                'db_cu': db_cu, 'db_gi_comercial': db_gi_comercial
            }
            print("[INFO] Bases cargadas exitosamente y guardadas en caché.")
        except Exception as e:
            print(f"[ERROR] Falló la carga de bases de datos: {str(e)}")
            return jsonify({"error": str(e)}), 500
    else:
        print("\n[INFO] Usando bases de datos desde caché. Generación rápida activada.")

    # 3. Extract Report Parameters
    nombre_predio = data.get('nombre_predio')
    lotcodigo_predio = data.get('lotcodigo_predio')
    coordenadas = tuple(data.get('coordenadas'))
    
    radio_gi = data.get('radio_gi')
    radio_cu = data.get('radio_cu')
    radio_comercio = data.get('radio_comercio')
    
    excluir_cu = data.get('excluir_codigos_cu', [])
    excluir_gi = data.get('excluir_codigos_gi_proyectos', [])
    excluir_gic = data.get('excluir_codigos_gi_comercio', [])

    # 4. Generate Document
    print(f"[INFO] Creando Word para {nombre_predio}...")
    try:
        crear_word(
            df_catastral=CACHE['df_catastral'],
            dict_dfs=CACHE['dict_dfs'],
            lotcodigo=lotcodigo_predio,
            nombre_predio=nombre_predio,
            columna_id='LotCodigo',
            coordenadas=coordenadas,
            radio_gi=radio_gi,
            radio_cu=radio_cu,
            radio_comercio=radio_comercio,
            excluir_codigos_gi_proyectos=excluir_gi,
            excluir_codigos_cu=excluir_cu,
            excluir_codigos_gi_comercio=excluir_gic
        )
    except Exception as e:
        print(f"[ERROR] Error creando el documento: {str(e)}")
        return jsonify({"error": str(e)}), 500

    # 5. Serve the generated file
    file_path = f"{nombre_predio}/Reporte_{nombre_predio}.docx"
    
    if os.path.exists(file_path):
        return send_file(
            file_path,
            as_attachment=True,
            download_name=f"Reporte_{nombre_predio}.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        return jsonify({"error": "No se pudo encontrar el archivo generado."}), 500

if __name__ == '__main__':
    # Run the local server
    print("\n=========================================")
    print("Servidor Iniciado. Abre esta URL en tu navegador:")
    print("http://127.0.0.1:5000")
    print("=========================================\n")
    def open_browser():
        webbrowser.open("http://127.0.0.1:5000")
    
    # Inicia el temporizador para abrir la pestaña
    threading.Timer(1.2, open_browser).start()
    
    # IMPORTANTE: debug=False evita que se abran dos pestañas al mismo tiempo
    app.run(debug=False, port=5000)