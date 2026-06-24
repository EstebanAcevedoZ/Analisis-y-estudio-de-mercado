import pandas as pd
import numpy as np
from geopy.distance import geodesic
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# Autor: Esteban Acevedo Z.
# Fecha: 2026
# Proyecto: Herramienta 1 de Análisis Inmobiliario por Proximidad usando Galería Inmobiliaria
# ==========================================

def obtener_puntos_circulo(lat, lon, radio_km, puntos=100):
    angulos = np.linspace(0, 2*np.pi, puntos)
    delta_lat = radio_km / 111.1
    delta_lon = radio_km / (111.1 * np.cos(np.deg2rad(lat)))
    lats = lat + delta_lat * np.sin(angulos)
    lons = lon + delta_lon * np.cos(angulos)
    return lats, lons

def procesar_con_fechas_y_mapa(archivo_path, lista_predios, lista_radios, col_agrupadora, operaciones, col_peso=None):
    print("Cargando datos de GI...")
    df_proyectos = pd.read_excel(archivo_path, sheet_name='Base Proyectos', engine='pyxlsb')
    df_inmuebles = pd.read_excel(archivo_path, sheet_name='Base Formulada', engine='pyxlsb')

    # Convertir Fecha Inicio a datetime inmediatamente
    #df_proyectos['Fecha Inicio'] = pd.to_datetime(df_proyectos['Fecha Inicio'], dayfirst=True, errors='coerce')

    df_proyectos = df_proyectos.dropna(subset=['Latitud (Coordenada Real)', 'Longitud (Coordenada Real)'])
    col_vivienda_proyectos = "Tipo de Vivienda ( Vis o No Vis )"

    resultados_filas = []
    todas_las_columnas = set() 
    proyectos_para_mapa = pd.DataFrame()
    
    # Columnas fijas al inicio
    cols_base = [("Nombre Predio", "ID"), ("Coordenada", "Lat, Lon"), ("Radio", "km"), ("Total Proyectos", "Global")]
    
    # Crear nombre combinado si hay múltiples predios
    if len(lista_predios) > 1:
        nombre_combinado = "_".join([p['nombre'] for p in lista_predios])
    else:
        nombre_combinado = lista_predios[0]['nombre']

    for predio in lista_predios:
        nombre = predio['nombre']
        coord_centro = (predio['lat'], predio['lon'])
        print(f"Procesando: {nombre}...")
        
        def calcular_distancia(fila):
            try: return geodesic(coord_centro, (fila['Latitud (Coordenada Real)'], fila['Longitud (Coordenada Real)'])).km
            except: return float('inf')
            
        df_proyectos['Distancia_Actual'] = df_proyectos.apply(calcular_distancia, axis=1)

        for radio_km in lista_radios:
            mask_cercanos = (df_proyectos['Distancia_Actual'] <= radio_km) & (df_proyectos['Activo'] == "Si")
            proyectos_cercanos = df_proyectos[mask_cercanos].copy()
            ids_proyectos = proyectos_cercanos['Codproyecto'].unique()
            
            if radio_km == max(lista_radios):
                proyectos_para_mapa = pd.concat([proyectos_para_mapa, proyectos_cercanos]).drop_duplicates(subset=['Codproyecto'])

            fila_datos = {
                ("Nombre Predio", "ID"): nombre,
                ("Coordenada", "Lat, Lon"): f"{coord_centro[0]}, {coord_centro[1]}",
                ("Radio", "km"): radio_km,
                ("Total Proyectos", "Global"): len(ids_proyectos)
            }

            if not proyectos_cercanos.empty:
                # --- NUEVO: Cálculo de fechas extrema ---
                f_min = proyectos_cercanos['Fecha Inicio'].min()
                f_max = proyectos_cercanos['Fecha Inicio'].max()
                #f_min.strftime('%d/%m/%Y') if pd.notnull(f_min) else "N/A"
                #f_max.strftime('%d/%m/%Y') if pd.notnull(f_max) else "N/A"
                # Guardamos como string para que el Excel no lo altere
                fila_datos[("Fecha Inicio", "Antigua")] = f_min
                fila_datos[("Fecha Inicio", "Reciente")] = f_max

                # Conteo Proyectos
                conteo = proyectos_cercanos[col_vivienda_proyectos].value_counts()
                for tipo_v, cant in conteo.items():
                    clave = (col_vivienda_proyectos, tipo_v)
                    fila_datos[clave] = cant
                    todas_las_columnas.add(clave)

                # Métricas Inmuebles
                inmuebles_f = df_inmuebles[df_inmuebles['Codproyecto'].isin(ids_proyectos)].copy()

                def aplicar_metricas(group):
                    #res = {'Cant. Inmuebles': len(group)}
                    res = {}
                    for col, ops in operaciones.items():
                        if col not in group.columns: continue
                        vals = pd.to_numeric(group[col], errors='coerce')
                        if 'promedio' in ops: res[f'Promedio {col}'] = vals.mean()
                        if 'suma' in ops: res[f'Suma {col}'] = vals.sum()
                        if 'ponderado' in ops and col_peso in group.columns:
                            w = pd.to_numeric(group[col_peso], errors='coerce').fillna(0)
                            res[f'Ponderado {col}'] = (vals.fillna(0) * w).sum() / w.sum() if w.sum() > 0 else 0
                    return pd.Series(res)

                if not inmuebles_f.empty:
                    res_agrupado = inmuebles_f.groupby(col_agrupadora).apply(aplicar_metricas, include_groups=False)
                    for m in res_agrupado.columns:
                        for t in res_agrupado.index:
                            clave = (m, t)
                            fila_datos[clave] = res_agrupado.loc[t, m]
                            todas_las_columnas.add(clave)

            resultados_filas.append(fila_datos)

    # --- GENERACIÓN EXCEL CON COLUMNAS AL FINAL ---
    # Ordenamos: Base + Dinámicas + Fechas al final
    columnas_ordenadas = cols_base + sorted(list(todas_las_columnas)) + [("Fecha Inicio", "Antigua"), ("Fecha Inicio", "Reciente")]
    
    df_final = pd.DataFrame([[c[0] for c in columnas_ordenadas], [c[1] for c in columnas_ordenadas]] + 
                           [[r.get(c, 0) for c in columnas_ordenadas] for r in resultados_filas])
    
    nombre_excel = f"Analisis_Predio_{nombre_combinado}_gi.xlsx"
    df_final.to_excel(nombre_excel, index=False, header=False)
    print(f"\nExcel '{nombre_excel}' generado exitosamente.")

    # --- MAPA ---
    print("Generando mapa...")
    fig = go.Figure()

    # 1. Proyectos
    fig.add_trace(go.Scattermap(
        lat=proyectos_para_mapa['Latitud (Coordenada Real)'],
        lon=proyectos_para_mapa['Longitud (Coordenada Real)'],
        mode='markers',
        marker=dict(size=8, color='blue', opacity=0.5),
        text=proyectos_para_mapa['Codproyecto'],
        name='Todos los Proyectos'
    ))

    colores = px.colors.qualitative.Vivid
    for i, predio in enumerate(lista_predios):
        color = colores[i % len(colores)]
        nombre = predio['nombre']
        lat, lon = predio['lat'], predio['lon']
        
        # Marcador del Predio
        fig.add_trace(go.Scattermap(
            lat=[lat], lon=[lon],
            mode='markers',
            marker=dict(size=18, color=color, symbol='circle'),
            name=f"CENTRO: {nombre}", # Nombre personalizado en leyenda
            text=f"Predio: {nombre}"
        ))

        # Radios Individuales (Sin legendgroup para que se puedan quitar uno a uno)
        for r in lista_radios:
            c_lats, c_lons = obtener_puntos_circulo(lat, lon, r)
            fig.add_trace(go.Scattermap(
                lat=c_lats, lon=c_lons,
                mode='lines',
                line=dict(width=2, color=color),
                # El nombre incluye el predio y el radio para control total
                name=f"{nombre} - {r}km",
                showlegend=True,
                hoverinfo='skip'
            ))

    fig.update_layout(mapbox_style="open-street-map",
        map=dict(center=dict(lat=lista_predios[0]['lat'], lon=lista_predios[0]['lon']), zoom=13),
        margin={"r":0,"t":40,"l":0,"b":0},
        title="Análisis por Predio y Radio Individual",
        height=850,
        legend=dict(groupclick="togglegroup")
    )
    fig.write_html(f"Mapa_Predio_{nombre_combinado}_gi.html")
    fig.show()

# --- CONFIGURACIÓN ---
config = {
    "archivo_path": "La Galería Inmobiliaria/Renobo_Bogotá_Bases_05-26.xlsb",
    #modificar a conveniencia
    "lista_predios": [
        {"nombre": "Sosiego", "lat": 4.574917958587793, "lon": -74.07973063953699},
        {"nombre": "Santa Cecilia", "lat": 4.607374225638947, "lon": -74.20154301205329},
        {"nombre": "Bochica", "lat": 4.596810500578444, "lon": -74.11497420988263},
        {"nombre": "UG2", "lat": 4.489719442585045, "lon": -74.10985072610995}
    ],

    "lista_radios": [1, 2, 3], #Modificar a conveniencia (km)
    "col_agrupadora": "Tipo VIS",
    "col_peso": "Oferta Total Inm.",
    "operaciones": {
        "Oferta Total Inm.": ["suma"],
        "Dmay26": ["suma"],
        "Pmay26": ["ponderado"],
        "$m2": ["ponderado"],
        "Ventas promedio mensual": ["promedio"], #promedio de unidades que vende un proyecto
        "Ventas promedio mensual": ["suma"], #promeido de unidades que se vende en la zona
        "Area": ["ponderado"],
        "Alc.": ["ponderado"],
        "Baños Completos": ["ponderado"]
    }
}

procesar_con_fechas_y_mapa(**config)