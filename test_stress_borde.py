import time
import io
import urllib.request
import urllib.parse
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from werkzeug.security import generate_password_hash
from app import app, validar_excel_tarifas

URL_VERCEL = "https://03sistemacotizadormvp.vercel.app/"  

def enviar_post(client, payload):
    if URL_VERCEL.strip():
        url = URL_VERCEL.rstrip("/") + "/"
        data = urllib.parse.urlencode(payload).encode("utf-8")
        req = urllib.request.Request(url, data=data, method="POST")
        with urllib.request.urlopen(req, timeout=10.0) as resp:
            return resp.status, resp.read().decode("utf-8")
    else:
        resp = client.post("/", data=payload)
        return resp.status_code, resp.data.decode("utf-8")

def ejecutar_suite_qa():
    modo = f"REMOTO VERCEL ({URL_VERCEL})" if URL_VERCEL.strip() else "LOCAL (Flask Test Client)"
    print("=" * 78)
    print(f" SUITE AUTOMATIZADA DE PRUEBAS DE BORDE Y STRESS - MODO: {modo}")
    print("=" * 78)

    client = app.test_client()

    casos_borde = [
        ("P05: Peso negativo (-15 Ton)", {"tipoOperacion": "Importación", "puertoOrigen": "Shanghai (China)", "puertoDestino": "San Antonio (Chile)", "tipoEquipo": "20 HC", "pesoCarga": "-15", "unidadPeso": "TON", "valorCarga": ""}, "menor al mínimo"),
        ("P06: Menos del mínimo (0.005 Ton / 5 Kg)", {"tipoOperacion": "Importación", "puertoOrigen": "Ningbo (China)", "puertoDestino": "Valparaíso (Chile)", "tipoEquipo": "20 HC", "pesoCarga": "5", "unidadPeso": "KG", "valorCarga": ""}, "menor al mínimo"),
        ("P07: Límite mínimo exacto (0.01 Ton)", {"tipoOperacion": "Importación", "puertoOrigen": "Shanghai (China)", "puertoDestino": "San Antonio (Chile)", "tipoEquipo": "20 HC", "pesoCarga": "0.01", "unidadPeso": "TON", "valorCarga": ""}, "1 Unidad"),
        ("P08: Borde exacto Romana 20' HC (23.00 Ton)", {"tipoOperacion": "Importación", "puertoOrigen": "Qingdao (China)", "puertoDestino": "San Antonio (Chile)", "tipoEquipo": "20 HC", "pesoCarga": "23.00", "unidadPeso": "TON", "valorCarga": ""}, "1 Unidad"),
        ("P09: Sobrepaso mínimo Romana 20' HC (23.01 Ton)", {"tipoOperacion": "Importación", "puertoOrigen": "Qingdao (China)", "puertoDestino": "San Antonio (Chile)", "tipoEquipo": "20 HC", "pesoCarga": "23.01", "unidadPeso": "TON", "valorCarga": ""}, "2 Unidades"),
        ("P10: Sobrepaso mínimo Romana 40' HC (25.01 Ton)", {"tipoOperacion": "Importación", "puertoOrigen": "Ho Chi Minh (Vietnam)", "puertoDestino": "Valparaíso (Chile)", "tipoEquipo": "40 HC", "pesoCarga": "25.01", "unidadPeso": "TON", "valorCarga": ""}, "2 Unidades"),
        ("P11: Límite máximo exacto (10,000.00 Ton)", {"tipoOperacion": "Importación", "puertoOrigen": "Busan (Corea del Sur)", "puertoDestino": "Valparaíso (Chile)", "tipoEquipo": "40 HC", "pesoCarga": "10000.00", "unidadPeso": "TON", "valorCarga": "5000000"}, "400 Unidades"),
        ("P12: Más del máximo (10,000.01 Ton)", {"tipoOperacion": "Importación", "puertoOrigen": "Busan (Corea del Sur)", "puertoDestino": "San Antonio (Chile)", "tipoEquipo": "40 HC", "pesoCarga": "10000.01", "unidadPeso": "TON", "valorCarga": ""}, "supera el límite máximo"),
        ("P13: Exportación a Perú + Valor Carga ($15,000 USD)", {"tipoOperacion": "Exportación", "puertoOrigen": "Valparaíso (Chile)", "puertoDestino": "Callao (Perú)", "tipoEquipo": "20 HC", "pesoCarga": "20", "unidadPeso": "TON", "valorCarga": "15000"}, "15,900.00"),
    ]

    print("\n[1] EJECUTANDO PRUEBAS DE BORDE (MÍNIMO, MÁXIMO Y FUERA DE RANGO):")
    for nombre, payload, texto_esperado in casos_borde:
        t0 = time.perf_counter()
        status, html = enviar_post(client, payload)
        dt_ms = (time.perf_counter() - t0) * 1000.0
        exito = (status == 200) and (texto_esperado in html)
        estado = "EXITOSO (HTTP 200 - App estable)" if exito else f"REVISAR (HTTP {status})"
        print(f"  - {nombre:<50} -> {estado} [{dt_ms:.2f} ms]")

    total_req = 50 if URL_VERCEL.strip() else 300
    print(f"\n[2] EJECUTANDO P14: STRESS DE CONCURRENCIA HTTP ({total_req} COTIZACIONES SIMULTÁNEAS):")
    def tarea_concurrente(i):
        t_ini = time.perf_counter()
        payload = {
            "tipoOperacion": "Importación",
            "puertoOrigen": "Shanghai (China)",
            "puertoDestino": "San Antonio (Chile)",
            "tipoEquipo": "40 HC",
            "pesoCarga": str(25 + (i % 75)),
            "unidadPeso": "TON",
            "valorCarga": "10000"
        }
        if URL_VERCEL.strip():
            st, _ = enviar_post(None, payload)
            assert st == 200
        else:
            with app.test_client() as c_local:
                r = c_local.post("/", data=payload)
                assert r.status_code == 200
        return (time.perf_counter() - t_ini) * 1000.0

    with ThreadPoolExecutor(max_workers=20) as executor:
        tiempos_ms = list(executor.map(tarea_concurrente, range(total_req)))

    serie_t = pd.Series(tiempos_ms)
    print(f"  - Peticiones procesadas : {len(serie_t)}/{total_req} (0% caídas)")
    print(f"  - Media: {serie_t.mean():.2f} ms | Mediana: {serie_t.median():.2f} ms | Mín: {serie_t.min():.2f} ms | Máx: {serie_t.max():.2f} ms | Std: {serie_t.std():.2f} ms")

    print("\n[3] EJECUTANDO P15: STRESS DE INGESTA EXCEL (2,500 FILAS EN MEMORIA):")
    filas_2500 = [{
        "Modalidad": "Importación",
        "Puerto Origen": f"Puerto_Stress_{i}",
        "Pais Origen": "China",
        "Puerto Destino": "San Antonio" if i % 2 == 0 else "Valparaíso",
        "Pais Destino": "Chile",
        "Tipo de Ruta": "Directo / Transbordo",
        "Tarifa 20' Min (US$)": 1600,
        "Tarifa 20' Max (US$)": 2800,
        "Tarifa 40' Min (US$)": 2300,
        "Tarifa 40' Max (US$)": 3800,
        "Transito Min (dias)": 25,
        "Transito Max (dias)": 32
    } for i in range(2500)]
    buffer_excel = io.BytesIO()
    pd.DataFrame(filas_2500).to_excel(buffer_excel, index=False, sheet_name="Tarifas Referencia")
    buffer_excel.seek(0)

    t0_excel = time.perf_counter()
    ok_excel, rutas_val, _ = validar_excel_tarifas(buffer_excel)
    dt_excel_ms = (time.perf_counter() - t0_excel) * 1000.0
    print(f"  - Validación Pandas/Openpyxl (2,500 filas): {ok_excel} | Rutas únicas: {len(rutas_val)} | Tiempo: {dt_excel_ms:.2f} ms")

    print("\n[4] EJECUTANDO P16: BENCHMARK DE HASHING SCRYPT (20 USUARIOS SEGUIDOS):")
    tiempos_hash = []
    for idx in range(20):
        th0 = time.perf_counter()
        _ = generate_password_hash(f"CintacSeguridad2026_{idx}!")
        tiempos_hash.append((time.perf_counter() - th0) * 1000.0)
    s_hash = pd.Series(tiempos_hash)
    print(f"  - Media Scrypt: {s_hash.mean():.2f} ms | Mediana: {s_hash.median():.2f} ms | Mín: {s_hash.min():.2f} ms | Máx: {s_hash.max():.2f} ms")
    print("=" * 78)

if __name__ == "__main__":
    ejecutar_suite_qa()