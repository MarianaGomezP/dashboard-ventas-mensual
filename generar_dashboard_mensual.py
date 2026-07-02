import json
import math
import os
import pandas as pd

EXCEL_PATH = 'Archivo Crudo_Datos Cierre Ventas 2026_actualizado.xlsx'
TRAFICO_PATH = 'trafico.xlsx'
PPTO_VM2_PATH = 'ppto vm2.xlsx'
OUTPUT_PATH = 'dashboard_ventas_mensuales.html'

MES_LABEL = {1:'ENE',2:'FEB',3:'MAR',4:'ABR',5:'MAY',6:'JUN',
             7:'JUL',8:'AGO',9:'SEP',10:'OCT',11:'NOV',12:'DIC'}
MES_NOMBRE = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',
              7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}


def _clean(val):
    if val is None:
        return None
    try:
        if math.isnan(val) or math.isinf(val):
            return None
    except (TypeError, ValueError):
        pass
    return val


def _clean_records(records):
    return [{k: _clean(v) for k, v in row.items()} for row in records]


def cargar_datos(path):
    df = pd.read_excel(path, sheet_name='Datos-Acum')
    df.columns = df.columns.str.strip()
    # Normalize column name with double space
    if 'INGRESO VARIABLE REAL  2026' in df.columns:
        df = df.rename(columns={'INGRESO VARIABLE REAL  2026': 'INGRESO VARIABLE REAL 2026'})
    return df


def cargar_trafico(path):
    df = pd.read_excel(path, sheet_name='Tráfico (Base de datos)')
    df.columns = df.columns.str.strip()
    cols = ['CECO', 'FECHA', 'AÑO', 'MES', 'TRÁFICO', 'Anterior (Calendario Fecha/Fecha)']
    df = df[[c for c in cols if c in df.columns]]
    df['FECHA'] = pd.to_datetime(df['FECHA'])
    df['dow'] = df['FECHA'].dt.dayofweek  # 0=lunes ... 6=domingo
    return df


def _build_trafico(df):
    d = df[df['AÑO'] == 2026].copy()
    # Clasificación propia lunes-jueves ("ES") / viernes-domingo ("FDS"),
    # distinta de la columna SEMANA del archivo (que marca viernes como entre semana).
    d['grupo_dia'] = d['dow'].apply(lambda x: 'ES' if x <= 3 else 'FDS')
    d = d.rename(columns={
        'CECO': 'ceco',
        'MES': 'mes',
        'TRÁFICO': 'trafico',
        'Anterior (Calendario Fecha/Fecha)': 'trafico_ant',
    })
    d = d[['ceco', 'mes', 'grupo_dia', 'trafico', 'trafico_ant']]
    return _clean_records(d.to_dict(orient='records'))


def cargar_ppto_vm2(path):
    df = pd.read_excel(path, sheet_name='BD_COMPLETA_NEW')
    df.columns = df.columns.str.strip()
    cols = ['Nombre Activo', 'Mes', 'Área', 'Venta Proy 2026', 'tipo de ppto mm o mt', 'Categoria', 'Marcas Foco']
    return df[[c for c in cols if c in df.columns]]


def _build_ppto_vm2(df):
    d = df.rename(columns={
        'Nombre Activo': 'ceco',
        'Mes': 'mes',
        'Área': 'area',
        'Venta Proy 2026': 'venta_proy',
        'tipo de ppto mm o mt': 'mm_mt',
        'Categoria': 'categoria',
        'Marcas Foco': 'marca_foco',
    })
    return _clean_records(d.to_dict(orient='records'))


def _build_resumen(df):
    g = df.groupby(['MES', 'NOMBRE ACTIVO'], as_index=False).agg(
        venta=('VALOR VENTA PREVALECE 2026', 'sum'),
        venta_ant=('VALOR VENTA PREVALECE 2025', 'sum'),
        ppto=('PRESUPUESTO DEFINITIVO VENTAS 2026', 'sum'),
        area=('AREA FINAL 2026', 'sum'),
    ).rename(columns={'MES': 'mes', 'NOMBRE ACTIVO': 'ceco'})
    return _clean_records(g.to_dict(orient='records'))


def _build_por_marca(df):
    cols = {
        'MES': 'mes',
        'NOMBRE ACTIVO': 'ceco',
        'NOMBRE MARCA': 'marca',
        'CATEGORÍA': 'categoria',
        'MM O MT': 'mm_mt',
        'MARCA FOCO': 'marca_foco',
        'VALOR VENTA PREVALECE 2026': 'venta',
        'VALOR VENTA PREVALECE 2025': 'venta_ant',
        'PRESUPUESTO DEFINITIVO VENTAS 2026': 'ppto',
        'AREA FINAL 2026': 'area',
        'MARCA ACTIVA': 'marca_activa',
        'INGRESO ADMIN MARCA 2026': 'ingreso_admin',
        'INGRESO FIJO MARCA 2026': 'ingreso_fijo',
        'INGRESO VARIABLE REAL 2026': 'ingreso_variable',
        'PPT VARIABLE DEFINITIVO 2026': 'ppt_variable',
    }
    result = df[[c for c in cols if c in df.columns]].rename(columns=cols)
    return _clean_records(result.to_dict(orient='records'))


def _build_por_categoria(df):
    g = df.groupby(['MES', 'NOMBRE ACTIVO', 'CATEGORÍA'], as_index=False).agg(
        venta=('VALOR VENTA PREVALECE 2026', 'sum'),
        venta_ant=('VALOR VENTA PREVALECE 2025', 'sum'),
        ppto=('PRESUPUESTO DEFINITIVO VENTAS 2026', 'sum'),
    ).rename(columns={'MES': 'mes', 'NOMBRE ACTIVO': 'ceco', 'CATEGORÍA': 'categoria'})
    return _clean_records(g.to_dict(orient='records'))


def _build_area_por_ceco(df):
    g = df.groupby(['NOMBRE ACTIVO', 'NOMBRE MARCA'])['AREA FINAL 2026'].max().reset_index()
    totals = g.groupby('NOMBRE ACTIVO')['AREA FINAL 2026'].sum()
    return {k: _clean(float(v)) for k, v in totals.items()}


def construir_D(path, trafico_path=None, ppto_vm2_path=None):
    df = cargar_datos(path)
    meses = sorted(df['MES'].dropna().unique().astype(int).tolist())
    cecos = sorted(df['NOMBRE ACTIVO'].dropna().unique().tolist())
    categorias = sorted(df['CATEGORÍA'].dropna().unique().tolist())

    D = {
        'cecos': cecos,
        'meses': meses,
        'mes_label': {str(m): MES_LABEL.get(m, str(m)) for m in meses},
        'mes_nombre': {str(m): MES_NOMBRE.get(m, str(m)) for m in meses},
        'ultimo_mes': int(df['MES'].max()),
        'categorias': categorias,
        'total_registros': int(len(df)),
        'resumen': _build_resumen(df),
        'por_marca': _build_por_marca(df),
        'por_categoria': _build_por_categoria(df),
        'area_por_ceco': _build_area_por_ceco(df),
        'trafico': _build_trafico(cargar_trafico(trafico_path)) if trafico_path else [],
        'ppto_vm2': _build_ppto_vm2(cargar_ppto_vm2(ppto_vm2_path)) if ppto_vm2_path else [],
    }
    return D


def render_html(D, output_path):
    # ensure_ascii=True convierte Ñ/Á/etc. a \uXXXX para evitar problemas en el script tag
    data_json = json.dumps(D, ensure_ascii=True, default=str)
    html = TEMPLATE.replace('/*DATA_PLACEHOLDER*/', data_json)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Dashboard generado: {output_path}')


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base, EXCEL_PATH)
    trafico_path = os.path.join(base, TRAFICO_PATH)
    ppto_vm2_path = os.path.join(base, PPTO_VM2_PATH)
    out  = os.path.join(base, OUTPUT_PATH)
    print(f'Cargando: {path}')
    print(f'Cargando: {trafico_path}')
    print(f'Cargando: {ppto_vm2_path}')
    D = construir_D(path, trafico_path, ppto_vm2_path)
    print(f"  Meses    : {D['meses']}")
    print(f"  CECOs    : {len(D['cecos'])} centros")
    print(f"  Registros: {D['total_registros']}")
    print(f"  Trafico  : {len(D['trafico'])} registros")
    print(f"  Ppto VM2 : {len(D['ppto_vm2'])} registros")
    render_html(D, out)


TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Gran Plaza — Dashboard Ventas Mensuales 2026</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{
  --dg:#0d5938; --ora:#ff7c00; --red:#b21e28; --yel:#ffd13f;
  --jade:#00ad68; --pic:#0099e1; --pur:#7c3aed; --teal:#0d9488;
  --bg:#ffffff; --surf:#f7f8fa; --surf2:#eef0f3;
  --txt:#111827; --txt2:#4b5563; --txt3:#9ca3af;
  --brd:#e5e7eb; --brd2:#d1d5db;
  --sans:'Sora',sans-serif; --mono:'DM Mono',monospace;
  --r:10px; --r2:14px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--sans);background:var(--bg);color:var(--txt);font-size:14px;line-height:1.5}
/* HEADER */
.hdr{position:sticky;top:0;z-index:100;background:#fff;border-bottom:3px solid var(--yel);
  padding:10px 20px;display:flex;flex-wrap:wrap;align-items:center;gap:10px;box-shadow:0 2px 8px rgba(0,0,0,.07)}
.hdr-logo{display:flex;align-items:center;gap:10px;margin-right:8px}
.hdr-logo span{font-size:15px;font-weight:800;color:var(--dg);letter-spacing:-.3px;white-space:nowrap}
.hdr-logo small{font-size:11px;color:var(--txt2);display:block;font-weight:500}
.hdr-sep{width:1px;height:32px;background:var(--brd2);flex-shrink:0}
.ctrl-group{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.ctrl-label{font-size:11px;font-weight:700;color:var(--txt3);text-transform:uppercase;letter-spacing:.5px;white-space:nowrap}
/* PILLS / BUTTONS */
.pill-row{display:flex;gap:4px;flex-wrap:wrap;align-items:center}
.pill{padding:5px 12px;border-radius:20px;border:1.5px solid var(--brd2);background:#fff;
  font-family:var(--sans);font-size:12px;font-weight:600;cursor:pointer;transition:.15s;color:var(--txt2);white-space:nowrap}
.pill:hover{border-color:var(--dg);color:var(--dg)}
.pill.active{background:var(--dg);border-color:var(--dg);color:#fff}
.pill.active-ora{background:var(--ora);border-color:var(--ora);color:#fff}
.pill-sm{padding:3px 9px;font-size:11px}
/* VIEW TABS */
.vtab-row{display:flex;gap:0;border:1.5px solid var(--brd2);border-radius:8px;overflow:hidden}
.vtab{padding:6px 16px;font-family:var(--sans);font-size:12px;font-weight:700;cursor:pointer;
  border:none;background:#fff;color:var(--txt2);transition:.15s;letter-spacing:.2px;white-space:nowrap}
.vtab.active{background:var(--dg);color:#fff}
/* CECO TABS */
.ctab-row{display:flex;gap:6px;flex-wrap:wrap;padding:12px 20px 0;border-bottom:2px solid var(--brd)}
.ctab{padding:7px 14px;border-radius:8px 8px 0 0;border:1.5px solid var(--brd);border-bottom:none;
  font-family:var(--sans);font-size:12px;font-weight:600;cursor:pointer;background:var(--surf);
  color:var(--txt2);transition:.15s;margin-bottom:-2px}
.ctab.active{background:#fff;border-color:var(--dg);color:var(--dg);border-bottom-color:#fff}
/* CATEGORIA DROPDOWN */
.cat-wrap{position:relative}
.cat-btn{padding:5px 12px;border-radius:20px;border:1.5px solid var(--brd2);background:#fff;
  font-family:var(--sans);font-size:12px;font-weight:600;cursor:pointer;color:var(--txt2);
  white-space:nowrap;transition:.15s;display:flex;align-items:center;gap:6px}
.cat-btn.has-filter{border-color:var(--pic);color:var(--pic);background:#e0f2fe}
.cat-btn:hover{border-color:var(--dg);color:var(--dg)}
.cat-menu{position:absolute;top:calc(100% + 6px);left:0;z-index:200;background:#fff;
  border:1.5px solid var(--brd2);border-radius:var(--r);box-shadow:0 8px 24px rgba(0,0,0,.12);
  min-width:240px;padding:8px 0;display:none}
.cat-menu.open{display:block}
.cat-menu-item{display:flex;align-items:center;gap:10px;padding:7px 14px;cursor:pointer;
  font-size:13px;font-weight:500;color:var(--txt);transition:.1s}
.cat-menu-item:hover{background:var(--surf)}
.cat-menu-item input[type=checkbox]{accent-color:var(--dg);width:15px;height:15px;cursor:pointer}
.cat-menu-sep{height:1px;background:var(--brd);margin:6px 0}
/* MAIN LAYOUT */
main{padding:16px 20px;max-width:1600px;margin:0 auto}
.section-title{font-size:13px;font-weight:800;color:var(--txt2);text-transform:uppercase;
  letter-spacing:.8px;margin:18px 0 10px;display:flex;align-items:center;gap:8px}
.section-title::before{content:'';display:block;width:4px;height:16px;border-radius:2px;background:var(--dg)}
/* KPI GRID */
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px;margin-bottom:4px}
.kpi{background:var(--surf);border:1.5px solid var(--brd);border-radius:var(--r2);padding:14px 16px;
  animation:kUp .35s ease both}
.kpi:nth-child(1){animation-delay:.03s}.kpi:nth-child(2){animation-delay:.06s}
.kpi:nth-child(3){animation-delay:.09s}.kpi:nth-child(4){animation-delay:.12s}
.kpi:nth-child(5){animation-delay:.15s}.kpi:nth-child(6){animation-delay:.18s}
.kpi:nth-child(7){animation-delay:.21s}.kpi:nth-child(8){animation-delay:.24s}
@keyframes kUp{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
.kpi-label{font-size:10.5px;font-weight:700;color:var(--txt3);text-transform:uppercase;
  letter-spacing:.5px;margin-bottom:6px;line-height:1.3}
.kpi-val{font-family:var(--mono);font-size:22px;font-weight:500;color:var(--txt);line-height:1.1}
.kpi-val.sm{font-size:17px}
.kpi-val.xs{font-size:14px}
.kpi-sub{font-size:11px;color:var(--txt2);margin-top:4px;display:flex;align-items:center;gap:5px;flex-wrap:wrap}
/* CHIPS */
.chip{display:inline-flex;align-items:center;gap:3px;padding:2px 8px;border-radius:10px;
  font-size:11px;font-weight:700;font-family:var(--mono)}
.chip-g{background:#d1fae5;color:#065f46}
.chip-y{background:#fef3c7;color:#92400e}
.chip-r{background:#fee2e2;color:#991b1b}
.chip-n{background:var(--surf2);color:var(--txt2)}
.chip-b{background:#dbeafe;color:#1e40af}
/* GRIDS */
.g2{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:12px}
.g64{display:grid;grid-template-columns:1.6fr 1fr;gap:12px;margin-bottom:12px}
.g46{display:grid;grid-template-columns:1fr 1.6fr;gap:12px;margin-bottom:12px}
.g1{display:grid;grid-template-columns:1fr;gap:12px;margin-bottom:12px}
@media(max-width:1100px){.g2,.g3,.g64,.g46{grid-template-columns:1fr}}
/* CARD */
.card{background:#fff;border:1.5px solid var(--brd);border-radius:var(--r2);padding:16px;
  min-height:80px;overflow:hidden}
.card-title{font-size:12px;font-weight:800;color:var(--txt2);text-transform:uppercase;
  letter-spacing:.6px;margin-bottom:12px;display:flex;align-items:center;justify-content:space-between}
.ch-wrap{position:relative;width:100%}
/* RANKING TABLE */
.rtbl{width:100%;border-collapse:collapse;font-size:13px}
.rtbl{white-space:nowrap}
.rtbl th{text-align:left;font-size:10px;font-weight:700;color:var(--txt3);text-transform:uppercase;
  letter-spacing:.3px;padding:7px 12px;border-bottom:1.5px solid var(--brd)}
.rtbl td{padding:9px 12px;border-bottom:1px solid var(--surf2);vertical-align:middle}
.rtbl tr:last-child td{border-bottom:none}
.rtbl tr:hover td{background:var(--surf)}
.rtbl .num{font-family:var(--mono);text-align:right}
.rtbl .ceco-name{font-weight:600;font-size:12px}
.rtbl th{cursor:pointer;user-select:none}
.rtbl th.sort-asc::after{content:' ▲';font-size:9px}
.rtbl th.sort-desc::after{content:' ▼';font-size:9px}
.bar-bg{background:var(--surf2);border-radius:3px;height:5px;width:80px;display:inline-block;vertical-align:middle}
.bar-fill{height:5px;border-radius:3px;background:var(--dg)}
/* TOP MARCAS */
.blist{display:flex;flex-direction:column;gap:5px}
.brow{display:flex;align-items:center;justify-content:space-between;padding:6px 10px;
  border-radius:8px;background:var(--surf);gap:8px}
.brow:hover{background:var(--surf2)}
.brow-left{display:flex;flex-direction:column;gap:2px;min-width:0}
.brow-name{font-size:12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:180px}
.brow-sub{font-size:10px;color:var(--txt3)}
.brow-right{display:flex;align-items:center;gap:8px;flex-shrink:0}
.brow-val{font-family:var(--mono);font-size:12px;font-weight:500;text-align:right}
.badge-foco{display:inline-block;padding:1px 6px;border-radius:6px;font-size:9px;font-weight:700;
  background:#fef3c7;color:#92400e;text-transform:uppercase;letter-spacing:.3px}
/* MARCA TABLE */
.search-wrap{margin-bottom:10px}
.search-input{width:100%;padding:8px 12px;border:1.5px solid var(--brd2);border-radius:8px;
  font-family:var(--sans);font-size:13px;outline:none;transition:.15s}
.search-input:focus{border-color:var(--dg)}
.mtbl-wrap{overflow-x:auto}
.mtbl{width:100%;border-collapse:collapse;font-size:12px;min-width:900px}
.mtbl th{text-align:left;font-size:10px;font-weight:700;color:var(--txt3);text-transform:uppercase;
  letter-spacing:.3px;padding:7px 8px;border-bottom:1.5px solid var(--brd);cursor:pointer;
  white-space:nowrap;user-select:none;background:#fff;position:sticky;top:0;z-index:1}
.mtbl th:hover{color:var(--dg)}
.mtbl th.sort-asc::after{content:' ▲'}
.mtbl th.sort-desc::after{content:' ▼'}
.mtbl td{padding:6px 8px;border-bottom:1px solid var(--surf2);vertical-align:middle}
.mtbl tr:last-child td{border-bottom:none}
.mtbl tr:hover td{background:var(--surf)}
.mtbl .num{font-family:var(--mono);text-align:right;white-space:nowrap}
.mtbl .marca-cell{font-weight:600;max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.mtbl-wrap{max-height:420px;overflow-y:auto}
/* FOOTER */
.footer{text-align:center;font-size:11px;color:var(--txt3);padding:14px 20px;
  border-top:1px solid var(--brd);margin-top:8px}
/* INGRESO SECTION HIGHLIGHT */
.kpi.kpi-ing{border-left:3px solid var(--pic)}
.kpi.kpi-te{border-left:3px solid var(--pur)}
/* HIDDEN */
[hidden]{display:none!important}
</style>
</head>
<body>

<!-- HEADER -->
<header class="hdr">
  <div class="hdr-logo">
    <span>Gran Plaza<br><small>Dashboard Ventas Mensuales 2026</small></span>
  </div>
  <div class="hdr-sep"></div>

  <div class="ctrl-group">
    <span class="ctrl-label">Vista</span>
    <div class="vtab-row">
      <button class="vtab active" id="btn-ger" onclick="setView('gerencia',this)">Gerencia General</button>
      <button class="vtab" id="btn-ceco" onclick="setView('ceco',this)">Por CECO</button>
    </div>
  </div>

  <div class="hdr-sep"></div>

  <div class="ctrl-group">
    <span class="ctrl-label">Modo</span>
    <div class="vtab-row">
      <button class="vtab active" id="btn-unico" onclick="setModo(false)">Mes único</button>
      <button class="vtab" id="btn-comp" onclick="setModo(true)">Acumular meses</button>
    </div>
  </div>

  <div class="ctrl-group">
    <span class="ctrl-label">Mes</span>
    <div class="pill-row" id="mes-pills"></div>
  </div>

  <div class="hdr-sep"></div>

  <div class="ctrl-group">
    <span class="ctrl-label">Tipo marcas</span>
    <div class="pill-row">
      <button class="pill pill-sm" id="btn-mt" onclick="setMmMt('MT')">MT — Totales</button>
      <button class="pill pill-sm" id="btn-mm" onclick="setMmMt('MM')">MM — Mismas</button>
    </div>
  </div>

  <div class="ctrl-group">
    <span class="ctrl-label">Categoría</span>
    <div class="cat-wrap">
      <button class="cat-btn" id="cat-btn" onclick="toggleCatMenu(event)">TODAS ▾</button>
      <div class="cat-menu" id="cat-menu"></div>
    </div>
  </div>

  <div class="ctrl-group">
    <span class="ctrl-label">Marca</span>
    <div class="cat-wrap">
      <button class="cat-btn" id="marca-fil-btn" onclick="toggleMarcaFil(event)">TODAS ▾</button>
      <div class="cat-menu" id="marca-fil-menu">
        <div class="cat-menu-item" style="padding:4px 8px">
          <input type="search" id="marca-fil-search" placeholder="Buscar marca…" style="width:100%;font-size:12px;border:1px solid var(--brd);border-radius:4px;padding:3px 6px" oninput="filterMarcaFil(this.value)">
        </div>
        <div class="cat-menu-sep"></div>
        <div class="cat-menu-item" onclick="toggleAllMarca()">
          <input type="checkbox" id="marca-fil-all" checked>
          <label for="marca-fil-all" style="cursor:pointer;font-weight:700">Todas las marcas</label>
        </div>
        <div class="cat-menu-sep"></div>
        <div id="marca-fil-list" style="max-height:220px;overflow-y:auto"></div>
      </div>
    </div>
  </div>

  <div class="ctrl-group">
    <span class="ctrl-label">Foco</span>
    <div class="pill-row">
      <button class="pill pill-sm active" id="btn-todo-foco" onclick="setFoco(false)">Todo</button>
      <button class="pill pill-sm" id="btn-solo-foco" onclick="setFoco(true)">Solo Foco</button>
    </div>
  </div>
</header>

<!-- CECO TABS (solo vista CECO) -->
<div id="ctab-row" class="ctab-row" hidden></div>

<!-- MAIN -->
<main>

  <!-- ===== VISTA GERENCIA ===== -->
  <div id="view-gerencia">

    <div class="section-title">Ventas</div>
    <div class="kpi-grid" id="kpis-ventas-g"></div>

    <div class="section-title">Ingresos por Marca</div>
    <div class="kpi-grid" id="kpis-ingresos-g"></div>

    <div class="section-title">Tráfico</div>
    <div class="kpi-grid" id="kpis-trafico-g"></div>

    <div class="g3" style="margin-top:14px">
      <div class="card">
        <div class="card-title">Ranking de Venta por m²</div>
        <div id="tbl-vm2" style="overflow-x:auto"></div>
      </div>
      <div class="card">
        <div class="card-title">Evolución Mensual — Ventas</div>
        <div class="ch-wrap" style="height:260px"><canvas id="ch-evol-g"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Ranking de CECOs</div>
        <div id="tbl-rank" style="overflow-x:auto"></div>
      </div>
    </div>

    <div class="g3">
      <div class="card">
        <div class="card-title">Participación por Categoría</div>
        <div class="ch-wrap" style="height:210px"><canvas id="ch-cat-pie"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Cumplimiento Presupuesto por CECO</div>
        <div class="ch-wrap" style="height:210px"><canvas id="ch-ppto-g"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Ingresos por CECO</div>
        <div class="ch-wrap" style="height:210px"><canvas id="ch-ing-g"></canvas></div>
      </div>
    </div>

    <div class="g2">
      <div class="card" style="margin-bottom:12px">
        <div class="card-title">Marcas Foco Reportando Venta — Top Crecimiento</div>
        <div id="top-crec" class="blist"></div>
      </div>
      <div class="card">
        <div class="card-title">Marcas Foco Reportando Venta — Top Caída</div>
        <div id="top-caida" class="blist"></div>
      </div>
    </div>

    <div class="g2">
      <div class="card">
        <div class="card-title">Comparativo entre CECOs — Evolución</div>
        <div class="ch-wrap" style="height:300px"><canvas id="ch-comp-g"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Comparativo entre CECOs — Evolución Venta / m²</div>
        <div class="ch-wrap" style="height:300px"><canvas id="ch-comp-vm2-g"></canvas></div>
      </div>
    </div>

  </div><!-- /view-gerencia -->

  <!-- ===== VISTA CECO ===== -->
  <div id="view-ceco" hidden>

    <div class="section-title" id="ceco-titulo">CECO</div>

    <div class="kpi-grid" id="kpis-ventas-c"></div>

    <div class="section-title">Ingresos por Marca</div>
    <div class="kpi-grid" id="kpis-ingresos-c"></div>

    <div class="section-title">Tráfico</div>
    <div class="kpi-grid" id="kpis-trafico-c"></div>

    <div class="g2" style="margin-top:14px">
      <div class="card">
        <div class="card-title">Evolución Mensual</div>
        <div class="ch-wrap" style="height:220px"><canvas id="ch-evol-c"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Ingresos del Período</div>
        <div class="ch-wrap" style="height:220px"><canvas id="ch-ing-c"></canvas></div>
      </div>
    </div>

    <div class="g2">
      <div class="card">
        <div class="card-title">Ventas por Categoría</div>
        <div class="ch-wrap" style="height:220px"><canvas id="ch-cat-c"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Evolución Ingresos Mensuales</div>
        <div class="ch-wrap" style="height:220px"><canvas id="ch-ing-evol-c"></canvas></div>
      </div>
    </div>

    <div class="g2">
      <div class="card">
        <div class="card-title">Marcas Foco Reportando Venta — Top Crecimiento</div>
        <div id="top-marcas-c" class="blist"></div>
      </div>
      <div class="card">
        <div class="card-title">Marcas Foco Reportando Venta — Top Caída</div>
        <div id="top-riesgo-c" class="blist"></div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Detalle de Marcas</div>
      <div class="search-wrap">
        <input class="search-input" id="marca-search" type="search" placeholder="Buscar marca...">
      </div>
      <div class="mtbl-wrap">
        <table class="mtbl" id="mtbl">
          <thead>
            <tr>
              <th onclick="sortMarcas('marca')">Marca</th>
              <th onclick="sortMarcas('categoria')">Categoría</th>
              <th class="num" onclick="sortMarcas('area')">Área m²</th>
              <th class="num" onclick="sortMarcas('venta')">Venta 2026</th>
              <th class="num" onclick="sortMarcas('vxm2')">Venta/m²</th>
              <th class="num" onclick="sortMarcas('venta_ant')">Venta 2025</th>
              <th class="num" onclick="sortMarcas('dvs25')">vs 2025</th>
              <th class="num" onclick="sortMarcas('ppto')">PPT</th>
              <th class="num" onclick="sortMarcas('dppto')">vs PPT</th>
              <th class="num" onclick="sortMarcas('ingreso_admin')">Ing. Admin</th>
              <th class="num" onclick="sortMarcas('ingreso_fijo')">Ing. Fijo</th>
              <th class="num" onclick="sortMarcas('ingreso_variable')">Ing. Var.</th>
              <th class="num" onclick="sortMarcas('ppt_variable')">PPT Var.</th>
              <th class="num" onclick="sortMarcas('dppt_var')">vs PPT Var.</th>
              <th class="num" onclick="sortMarcas('te')">TE</th>
              <th style="text-align:center">Foco</th>
            </tr>
          </thead>
          <tbody id="mtbl-body"></tbody>
        </table>
      </div>
    </div>

  </div><!-- /view-ceco -->

</main>

<footer class="footer" id="footer-info">
  Dashboard Ventas Mensuales 2026 — Gran Plaza
</footer>

<script>
const D = /*DATA_PLACEHOLDER*/;

/* ── ESTADO ── */
const MES_SHORT = D.mes_label;
const MES_LONG  = D.mes_nombre;
const CC = ['#0099e1','#ffd13f','#00ad68','#ff7c00','#b21e28','#7c3aed','#0d9488','#f59e0b'];

let mesesActivos = [D.ultimo_mes];
let modoComparar = false;
let view = 'gerencia';
let ceco = D.cecos[0];
let mmMt = 'MT';
let catFiltro = [];
let soloFoco = false;
let charts = {};
let mSort = {col:'venta', dir:-1};
let mSearch = '';
let marcaFiltro = [];
let rSort = {col:'v', dir:-1};

/* ── FILTROS ── */
function inMes(r)   { return mesesActivos.includes(r.mes); }
function inMmMt(r)  { return mmMt === 'MT' || r.mm_mt === 'MM'; }
function inCat(r)   { return catFiltro.length === 0 || catFiltro.includes(r.categoria); }
function inFoco(r)  { return !soloFoco || r.marca_foco === 'SI'; }
function inMarca(r) { return marcaFiltro.length === 0 || marcaFiltro.includes(r.marca); }

/* ── FILTROS PRESUPUESTO VM2 ── */
function inMesPpto(r)  { return mesesActivos.includes(r.mes); }
function inMmMtPpto(r) { return mmMt === 'MT' || r.mm_mt === 'MM'; }
function filteredPptoVm2(extra) {
  return D.ppto_vm2.filter(r => inMesPpto(r) && inMmMtPpto(r) && inCat(r) && inFoco(r) && (!extra || extra(r)));
}
function filteredPptoVm2Ceco(c) { return filteredPptoVm2(r => r.ceco === c); }

function calcPptoVm2(rows) {
  const venta = sum(rows, 'venta_proy');
  const area = sum(rows, 'area');
  return area > 0 ? venta / area : null;
}

function filtered(extra) {
  return D.por_marca.filter(r => inMes(r) && inMmMt(r) && inCat(r) && inFoco(r) && inMarca(r) && (!extra || extra(r)));
}
function filteredCeco(c) { return filtered(r => r.ceco === c); }

/* ── FILTROS TRÁFICO ── */
function inMesT(r) { return mesesActivos.includes(r.mes); }
function filteredTrafico(extra) {
  return D.trafico.filter(r => inMesT(r) && (!extra || extra(r)));
}
function filteredTraficoCeco(c) { return filteredTrafico(r => r.ceco === c); }

/* Por todos los meses (para gráficos de evolución) */
function filteredAllMeses(extra) {
  return D.por_marca.filter(r => inMmMt(r) && inCat(r) && inFoco(r) && (!extra || extra(r)));
}

/* ── AGRUPACIÓN DE MARCAS MULTI-MES ── */
function groupBy(rows, keyFn) {
  const m = new Map();
  for (const r of rows) {
    const k = keyFn(r);
    if (!m.has(k)) m.set(k, {...r, venta:0, venta_ant:0, ppto:0, area:0,
      ingreso_admin:0, ingreso_fijo:0, ingreso_variable:0, ppt_variable:0});
    const g = m.get(k);
    g.venta += r.venta||0; g.venta_ant += r.venta_ant||0; g.ppto += r.ppto||0;
    g.area = Math.max(g.area||0, r.area||0);
    g.ingreso_admin += r.ingreso_admin||0; g.ingreso_fijo += r.ingreso_fijo||0;
    g.ingreso_variable += r.ingreso_variable||0; g.ppt_variable += r.ppt_variable||0;
  }
  return [...m.values()];
}

/* ── HELPERS NUMÉRICOS ── */
function sum(rows, f) { return rows.reduce((a,r) => a+(r[f]||0), 0); }
// Para MOSTRAR área: única por marca (no acumula meses), solo marcas activas
function areaActiva(rows) {
  const seen = new Map();
  rows.filter(r => r.marca_activa === 'SI').forEach(r => {
    const k = r.ceco + '|' + r.marca;
    seen.set(k, Math.max(seen.get(k)||0, r.area||0));
  });
  return [...seen.values()].reduce((a,v) => a+v, 0);
}
// Para CALCULAR VM2 e Ing/M2: suma acumulada por mes, solo marcas activas
function areaActivaCalc(rows) {
  return sum(rows.filter(r => r.marca_activa === 'SI'), 'area');
}
function delta(a, b)  { return b ? (a/b - 1)*100 : null; }
function pctOf(a, b)  { return b ? (a/b)*100 : null; }
function safe(v)      { return (v === null || v === undefined || isNaN(v)) ? null : v; }

/* ── FORMATO ── */
const COP = new Intl.NumberFormat('es-CO', {maximumFractionDigits:0});
function fv(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const a = Math.abs(n);
  if (a >= 1e9) return '$' + (n/1e9).toFixed(1) + 'MM';
  if (a >= 1e6) return '$' + (n/1e6).toFixed(1) + 'M';
  if (a >= 1e3) return '$' + (n/1e3).toFixed(0) + 'K';
  return '$' + COP.format(n);
}
function fvFull(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return '$' + COP.format(Math.round(n));
}
function fn(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return COP.format(Math.round(n));
}
function ftraf(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const a = Math.abs(n);
  if (a >= 1e6) return (n/1e6).toFixed(1) + 'M';
  if (a >= 1e3) return (n/1e3).toFixed(0) + 'K';
  return COP.format(Math.round(n));
}
function fp(v, dec=1) {
  if (v === null || v === undefined || isNaN(v)) return '—';
  return (v >= 0 ? '+' : '') + v.toFixed(dec) + '%';
}
function fm2(n) {
  if (!n || isNaN(n)) return '—';
  return '$' + COP.format(Math.round(n)) + '/m²';
}
function fpt(v) { // porcentaje sin signo
  if (v === null || v === undefined || isNaN(v)) return '—';
  return v.toFixed(1) + '%';
}

/* ── CHIPS ── */
function chip(val) {
  if (val === null || val === undefined || isNaN(val)) return '<span class="chip chip-n">—</span>';
  const cls = val >= 0 ? 'chip-g' : val >= -5 ? 'chip-y' : 'chip-r';
  const icon = val >= 0 ? '▲' : '▼';
  return `<span class="chip ${cls}">${icon} ${Math.abs(val).toFixed(1)}%</span>`;
}
function chipPpto(val) {
  if (val === null || val === undefined || isNaN(val)) return '<span class="chip chip-n">—</span>';
  const cls = val >= 100 ? 'chip-g' : val >= 90 ? 'chip-y' : 'chip-r';
  const icon = val >= 100 ? '▲' : '▼';
  return `<span class="chip ${cls}">${icon} ${val.toFixed(1)}%</span>`;
}
function chipTE(val) {
  // TE: ratio ingreso/venta — typical range 5-15%, lower is better
  if (val === null || val === undefined || isNaN(val)) return '<span class="chip chip-n">—</span>';
  const cls = val <= 6 ? 'chip-g' : val <= 10 ? 'chip-y' : 'chip-r';
  return `<span class="chip ${cls}">${val.toFixed(2)}%</span>`;
}

/* ── CECO LABEL ── */
function cecoShort(c) { return c.replace('GRAN PLAZA ',''); }

/* ── DESTRUIR CHART ── */
function dc(id) { if (charts[id]) { charts[id].destroy(); delete charts[id]; } }

/* ── LABEL MES ACTIVO ── */
function mesLabel() {
  if (mesesActivos.length === 1) return MES_LONG[mesesActivos[0]] || mesesActivos[0];
  return mesesActivos.map(m => MES_SHORT[m]).join('+');
}

/* ════════════════════════════════════
   INICIALIZACIÓN
════════════════════════════════════ */
function init() {
  buildMesPills();
  buildCecoTabs();
  buildCatMenu();
  buildMarcaFil();
  updateMtBtn();
  document.getElementById('footer-info').textContent =
    `${D.total_registros} registros · Meses: ${D.meses.map(m=>MES_SHORT[m]).join(', ')} · Gran Plaza 2026`;
  renderAll();
  document.addEventListener('click', e => {
    if (!document.getElementById('cat-btn').contains(e.target) &&
        !document.getElementById('cat-menu').contains(e.target)) {
      document.getElementById('cat-menu').classList.remove('open');
    }
    if (!document.getElementById('marca-fil-btn').contains(e.target) &&
        !document.getElementById('marca-fil-menu').contains(e.target)) {
      document.getElementById('marca-fil-menu').classList.remove('open');
    }
  });
  document.getElementById('marca-search').addEventListener('input', e => {
    mSearch = e.target.value.toLowerCase();
    renderMarcaTable();
  });
}

function buildMesPills() {
  const row = document.getElementById('mes-pills');
  row.innerHTML = '';
  D.meses.forEach(m => {
    const b = document.createElement('button');
    b.className = 'pill pill-sm' + (mesesActivos.includes(m) ? ' active' : '');
    b.id = 'pill-mes-' + m;
    b.textContent = MES_SHORT[m] || m;
    b.onclick = () => toggleMes(m);
    row.appendChild(b);
  });
}

function buildCecoTabs() {
  const row = document.getElementById('ctab-row');
  row.innerHTML = '';
  D.cecos.forEach(c => {
    const b = document.createElement('button');
    b.className = 'ctab' + (c === ceco ? ' active' : '');
    b.id = 'ctab-' + c;
    b.textContent = cecoShort(c);
    b.onclick = () => setCeco(c);
    row.appendChild(b);
  });
}

function buildCatMenu() {
  const menu = document.getElementById('cat-menu');
  menu.innerHTML = `
    <div class="cat-menu-item" onclick="toggleAllCat()">
      <input type="checkbox" id="cat-all" checked> <label for="cat-all" style="cursor:pointer;font-weight:700">Todas las categorías</label>
    </div>
    <div class="cat-menu-sep"></div>
  `;
  D.categorias.forEach(cat => {
    const id = 'cat-' + cat.replace(/\s/g,'_');
    const div = document.createElement('div');
    div.className = 'cat-menu-item';
    div.innerHTML = `<input type="checkbox" id="${id}" value="${cat}" onchange="onCatChange()"> <label for="${id}" style="cursor:pointer">${cat}</label>`;
    menu.appendChild(div);
  });
}

function toggleCatMenu(e) {
  e.stopPropagation();
  document.getElementById('cat-menu').classList.toggle('open');
}

function toggleAllCat() {
  catFiltro = [];
  document.querySelectorAll('#cat-menu input[type=checkbox]').forEach(cb => cb.checked = true);
  updateCatBtn();
  renderAll();
}

function onCatChange() {
  const checked = [...document.querySelectorAll('#cat-menu input[value]')].filter(c=>c.checked).map(c=>c.value);
  catFiltro = checked.length === D.categorias.length ? [] : checked;
  document.getElementById('cat-all').checked = (catFiltro.length === 0);
  updateCatBtn();
  renderAll();
}

function updateCatBtn() {
  const btn = document.getElementById('cat-btn');
  if (catFiltro.length === 0) { btn.textContent = 'TODAS ▾'; btn.classList.remove('has-filter'); }
  else { btn.textContent = catFiltro.length + ' CAT. ▾'; btn.classList.add('has-filter'); }
}

function buildMarcaFil() {
  const rows = (view === 'ceco') ? D.por_marca.filter(r => r.ceco === ceco) : D.por_marca;
  const marcas = [...new Set(rows.map(r => r.marca))].sort();
  marcaFiltro = marcaFiltro.filter(m => marcas.includes(m));
  const list = document.getElementById('marca-fil-list');
  list.innerHTML = '';
  marcas.forEach(m => {
    const id = 'mf-' + m.replace(/[^a-zA-Z0-9]/g,'_');
    const div = document.createElement('div');
    div.className = 'cat-menu-item marca-fil-item';
    div.dataset.marca = m.toLowerCase();
    div.innerHTML = `<input type="checkbox" id="${id}" value="${m}" onchange="onMarcaFilChange()"> <label for="${id}" style="cursor:pointer;font-size:12px">${m}</label>`;
    list.appendChild(div);
  });
  document.getElementById('marca-fil-all').checked = (marcaFiltro.length === 0);
  updateMarcaFilBtn();
}

function toggleMarcaFil(e) {
  e.stopPropagation();
  document.getElementById('marca-fil-menu').classList.toggle('open');
}

function filterMarcaFil(q) {
  const s = q.toLowerCase();
  document.querySelectorAll('.marca-fil-item').forEach(d => {
    d.style.display = d.dataset.marca.includes(s) ? '' : 'none';
  });
}

function toggleAllMarca() {
  const all = document.getElementById('marca-fil-all').checked;
  document.querySelectorAll('#marca-fil-list input[type=checkbox]').forEach(c => { c.checked = all; });
  marcaFiltro = [];
  updateMarcaFilBtn();
  renderAll();
}

function onMarcaFilChange() {
  const total = document.querySelectorAll('#marca-fil-list input[value]').length;
  const checked = [...document.querySelectorAll('#marca-fil-list input[value]')].filter(c => c.checked).map(c => c.value);
  marcaFiltro = checked.length === total ? [] : checked;
  document.getElementById('marca-fil-all').checked = (marcaFiltro.length === 0);
  updateMarcaFilBtn();
  renderAll();
}

function updateMarcaFilBtn() {
  const btn = document.getElementById('marca-fil-btn');
  if (marcaFiltro.length === 0) { btn.textContent = 'TODAS ▾'; btn.classList.remove('has-filter'); }
  else { btn.textContent = marcaFiltro.length + ' selec. ▾'; btn.classList.add('has-filter'); }
}

/* ── CONTROLES ── */
function toggleMes(m) {
  if (modoComparar) {
    const idx = mesesActivos.indexOf(m);
    if (idx >= 0 && mesesActivos.length > 1) mesesActivos.splice(idx, 1);
    else if (idx < 0) mesesActivos.push(m);
  } else {
    mesesActivos = [m];
  }
  buildMesPills();
  renderAll();
}

function setModo(comp) {
  modoComparar = comp;
  if (!comp) mesesActivos = [Math.max(...mesesActivos)];
  document.getElementById('btn-unico').classList.toggle('active', !comp);
  document.getElementById('btn-comp').classList.toggle('active', comp);
  buildMesPills();
  renderAll();
}

function setMmMt(val) {
  mmMt = val;
  updateMtBtn();
  renderAll();
}

function updateMtBtn() {
  document.getElementById('btn-mt').classList.toggle('active', mmMt === 'MT');
  document.getElementById('btn-mm').classList.toggle('active', mmMt === 'MM');
}

function setFoco(only) {
  soloFoco = only;
  document.getElementById('btn-todo-foco').classList.toggle('active', !only);
  document.getElementById('btn-solo-foco').classList.toggle('active', only);
  renderAll();
}

function setView(v, el) {
  view = v;
  document.getElementById('view-gerencia').hidden = (v !== 'gerencia');
  document.getElementById('view-ceco').hidden = (v !== 'ceco');
  document.getElementById('ctab-row').hidden = (v !== 'ceco');
  document.getElementById('btn-ger').classList.toggle('active', v==='gerencia');
  document.getElementById('btn-ceco').classList.toggle('active', v==='ceco');
  buildMarcaFil();
  renderAll();
}

function setCeco(c) {
  ceco = c;
  document.querySelectorAll('.ctab').forEach(b => b.classList.remove('active'));
  const tab = document.getElementById('ctab-' + c);
  if (tab) tab.classList.add('active');
  buildMarcaFil();
  renderCecoView();
}

/* ── RENDER PRINCIPAL ── */
function renderAll() {
  if (view === 'gerencia') renderGerencia();
  else renderCecoView();
}

/* ════════════════════════════════════
   VISTA GERENCIA
════════════════════════════════════ */
function renderGerencia() {
  renderKpisVentasG();
  renderKpisIngresosG();
  renderKpisTraficoG();
  renderEvolG();
  renderRankingG();
  renderVM2G();
  renderCatPieG();
  renderPptoG();
  renderIngresosG();
  renderTopMarcasG();
  renderComparativoG();
  renderComparativoVM2G();
}

function renderVM2G() {
  const el = document.getElementById('tbl-vm2');
  const data = D.cecos.map(c => {
    const rows = filtered(r => r.ceco === c);
    const v = sum(rows, 'venta');
    const ac = areaActivaCalc(rows);
    const vxm2 = ac > 0 ? v / ac : null;
    const pptoVm2 = calcPptoVm2(filteredPptoVm2Ceco(c));
    return {c, vxm2, cumpl: pctOf(vxm2, pptoVm2)};
  }).filter(d => d.vxm2 !== null).sort((a, b) => b.vxm2 - a.vxm2);
  el.innerHTML = `<table class="rtbl"><thead><tr>
    <th>CECO</th><th class="num">Venta / m²</th><th class="num">% Cumpl. PPTO VM2</th>
  </tr></thead><tbody>` +
  data.map(d => `<tr>
    <td class="ceco-name">${cecoShort(d.c)}</td>
    <td class="num">${fm2(d.vxm2)}</td>
    <td class="num">${chipPpto(d.cumpl)}</td>
  </tr>`).join('') + '</tbody></table>';
}

function renderKpisVentasG() {
  const rows = filtered();
  const venta = sum(rows,'venta');
  const venta_ant = sum(rows,'venta_ant');
  const ppto = sum(rows,'ppto');
  const area = areaActiva(rows);          // para mostrar
  const areaCalc = areaActivaCalc(rows);  // para calcular VM2
  const dvsAnt = delta(venta, venta_ant);
  const dvsPpto = pctOf(venta, ppto);
  const vxm2 = areaCalc > 0 ? venta/areaCalc : null;
  const pptoVm2 = calcPptoVm2(filteredPptoVm2());

  const marcasRep = new Set(rows.filter(r=>r.venta>0).map(r=>r.ceco+'|'+r.marca)).size;
  const marcasFoco = new Set(rows.filter(r=>r.venta>0 && r.marca_foco==='SI').map(r=>r.ceco+'|'+r.marca)).size;

  // Mejor y peor CECO por % vs PPT
  const pptByCeco = {};
  const ventaByCeco = {};
  D.cecos.forEach(c => {
    const cr = filtered(r=>r.ceco===c);
    const v = sum(cr,'venta'); const p = sum(cr,'ppto');
    ventaByCeco[c] = v; pptByCeco[c] = p;
  });
  const ratios = D.cecos.map(c => ({c, r: pctOf(ventaByCeco[c], pptByCeco[c])})).filter(x=>x.r!=null);
  ratios.sort((a,b)=>b.r-a.r);
  const mejorCeco = ratios[0]; const peorCeco = ratios[ratios.length-1];

  document.getElementById('kpis-ventas-g').innerHTML = `
    <div class="kpi">
      <div class="kpi-label">Venta Total 2026</div>
      <div class="kpi-val sm">${fv(venta)}</div>
      <div class="kpi-sub">${fvFull(venta)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">vs Año Anterior 2025</div>
      <div class="kpi-val sm">${chip(dvsAnt)}</div>
      <div class="kpi-sub">2025: ${fv(venta_ant)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">vs Presupuesto</div>
      <div class="kpi-val sm">${chipPpto(dvsPpto)}</div>
      <div class="kpi-sub">PPT: ${fv(ppto)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Venta / m²</div>
      <div class="kpi-val xs">${fm2(vxm2)}</div>
      <div class="kpi-sub">Área: ${fn(area)} m²</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Presupuesto VM2</div>
      <div class="kpi-val xs">${fm2(pptoVm2)}</div>
      <div class="kpi-sub">${chipPpto(pctOf(vxm2, pptoVm2))} cumpl.</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Marcas Reportando</div>
      <div class="kpi-val">${marcasRep}</div>
      <div class="kpi-sub">con venta &gt; 0</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Marcas Foco</div>
      <div class="kpi-val">${marcasFoco}</div>
      <div class="kpi-sub">con venta &gt; 0</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Mejor CECO (PPT)</div>
      <div class="kpi-val xs">${mejorCeco ? cecoShort(mejorCeco.c) : '—'}</div>
      <div class="kpi-sub">${mejorCeco ? chipPpto(mejorCeco.r) : ''}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">CECO a Reforzar</div>
      <div class="kpi-val xs">${peorCeco ? cecoShort(peorCeco.c) : '—'}</div>
      <div class="kpi-sub">${peorCeco ? chipPpto(peorCeco.r) : ''}</div>
    </div>
  `;
}

/* ── TRÁFICO ── */
function calcTraficoKpis(rows) {
  const total = sum(rows, 'trafico');
  const totalAnt = sum(rows, 'trafico_ant');
  const variacion = delta(total, totalAnt);
  const es = rows.filter(r => r.grupo_dia === 'ES');
  const fds = rows.filter(r => r.grupo_dia === 'FDS');
  const promEs  = es.length  ? sum(es, 'trafico') / es.length : null;
  const promFds = fds.length ? sum(fds, 'trafico') / fds.length : null;
  return { total, totalAnt, variacion, promEs, promFds };
}

function kpisTraficoHTML(k) {
  return `
    <div class="kpi">
      <div class="kpi-label">Tráfico Total 2026</div>
      <div class="kpi-val sm">${ftraf(k.total)}</div>
      <div class="kpi-sub">${fn(k.total)} personas</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Tráfico Total Año Anterior</div>
      <div class="kpi-val sm">${ftraf(k.totalAnt)}</div>
      <div class="kpi-sub">${fn(k.totalAnt)} personas</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">% Variación vs Año Anterior</div>
      <div class="kpi-val sm">${chip(k.variacion)}</div>
      <div class="kpi-sub">2026 vs 2025</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Tráfico Prom. Entre Semana</div>
      <div class="kpi-val sm">${ftraf(k.promEs)}</div>
      <div class="kpi-sub">Lunes a jueves</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Tráfico Prom. Fin de Semana</div>
      <div class="kpi-val sm">${ftraf(k.promFds)}</div>
      <div class="kpi-sub">Viernes a domingo</div>
    </div>
  `;
}

function renderKpisTraficoG() {
  const k = calcTraficoKpis(filteredTrafico());
  document.getElementById('kpis-trafico-g').innerHTML = kpisTraficoHTML(k);
}

function renderKpisTraficoC() {
  const k = calcTraficoKpis(filteredTraficoCeco(ceco));
  document.getElementById('kpis-trafico-c').innerHTML = kpisTraficoHTML(k);
}

function renderKpisIngresosG() {
  const rows = filtered();
  const admin = sum(rows,'ingreso_admin');
  const fijo  = sum(rows,'ingreso_fijo');
  const varR  = sum(rows,'ingreso_variable');
  const pptV  = sum(rows,'ppt_variable');
  const area     = areaActiva(rows);      // para mostrar
  const areaCalc = areaActivaCalc(rows);  // para calcular /m²
  const venta = sum(rows,'venta');
  const total = admin + fijo + varR;
  const te    = venta > 0 ? (total/venta)*100 : null;
  const dvV   = pctOf(varR, pptV);

  document.getElementById('kpis-ingresos-g').innerHTML = `
    <div class="kpi kpi-ing">
      <div class="kpi-label">Ingreso Total</div>
      <div class="kpi-val sm">${fv(total)}</div>
      <div class="kpi-sub">Admin+Fijo+Variable</div>
    </div>
    <div class="kpi kpi-te">
      <div class="kpi-label">Tasa de Esfuerzo (TE)</div>
      <div class="kpi-val sm">${chipTE(te)}</div>
      <div class="kpi-sub">Ing.Total / Venta 2026</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">Ingreso Admin</div>
      <div class="kpi-val sm">${fv(admin)}</div>
      <div class="kpi-sub">${fm2(areaCalc>0?admin/areaCalc:null)}</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">Ingreso Fijo</div>
      <div class="kpi-val sm">${fv(fijo)}</div>
      <div class="kpi-sub">${fm2(areaCalc>0?fijo/areaCalc:null)}</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">Ing. Variable Real</div>
      <div class="kpi-val sm">${fv(varR)}</div>
      <div class="kpi-sub">${fm2(areaCalc>0?varR/areaCalc:null)}</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">PPT Variable</div>
      <div class="kpi-val sm">${fv(pptV)}</div>
      <div class="kpi-sub">Meta ingreso variable</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">Cumpl. Variable</div>
      <div class="kpi-val sm">${chipPpto(dvV)}</div>
      <div class="kpi-sub">Variable vs PPT</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">Ingreso Total / m²</div>
      <div class="kpi-val xs">${fm2(areaCalc>0?total/areaCalc:null)}</div>
      <div class="kpi-sub">Área: ${fn(area)} m²</div>
    </div>
  `;
}

function renderEvolG() {
  dc('ch-evol-g');
  const allRows = filteredAllMeses();
  const labels = D.meses.map(m => MES_SHORT[m]);
  const venta   = D.meses.map(m => sum(allRows.filter(r=>r.mes===m), 'venta'));
  const ventaA  = D.meses.map(m => sum(allRows.filter(r=>r.mes===m), 'venta_ant'));
  const ppto    = D.meses.map(m => sum(allRows.filter(r=>r.mes===m), 'ppto'));
  const marcasXMes = D.meses.map(m =>
    new Set(allRows.filter(r=>r.mes===m && r.venta>0).map(r=>r.ceco+'|'+r.marca)).size
  );
  const varPctG = D.meses.map((_,i) => ventaA[i] ? (venta[i] - ventaA[i]) / ventaA[i] * 100 : null);
  const ctx = document.getElementById('ch-evol-g').getContext('2d');
  charts['ch-evol-g'] = new Chart(ctx, {
    type:'line',
    data:{labels, datasets:[
      {label:'Venta 2026', data:venta, borderColor:CC[0], backgroundColor:CC[0]+'22',
       tension:.35, fill:true, pointRadius:4, borderWidth:2},
      {label:'Venta 2025', data:ventaA, borderColor:CC[2], borderDash:[5,3],
       tension:.35, pointRadius:3, borderWidth:1.5},
      {label:'Presupuesto', data:ppto, borderColor:CC[3], borderDash:[3,3],
       tension:.35, pointRadius:3, borderWidth:1.5},
    ]},
    options:{responsive:true, maintainAspectRatio:false,
      plugins:{legend:{labels:{font:{family:'Sora',size:11},boxWidth:14}},
        tooltip:{mode:'index', intersect:false, callbacks:{
          label: c => '  ' + c.dataset.label + ': ' + fvFull(c.raw),
          footer: items => {
            if (!items.length) return '';
            const idx = items[0].dataIndex;
            const v = varPctG[idx];
            const pctStr = v !== null ? (v >= 0 ? '+' : '') + v.toFixed(1) + '% 2026 vs 2025' : '';
            return ['Marcas reportando: ' + marcasXMes[idx], pctStr].filter(Boolean).join('\n');
          }
        }}},
      scales:{y:{ticks:{font:{family:'DM Mono',size:10},
        callback:v=>fv(v)}, grid:{color:'#f0f0f0'}},
        x:{ticks:{font:{family:'Sora',size:11}}}}}
  });
}

function sortRanking(col) {
  if (rSort.col === col) rSort.dir *= -1;
  else { rSort.col = col; rSort.dir = -1; }
  renderRankingG();
}

function renderRankingG() {
  const el = document.getElementById('tbl-rank');
  const data = D.cecos.map(c => {
    const rows = filtered(r=>r.ceco===c);
    const v = sum(rows,'venta'); const va = sum(rows,'venta_ant'); const p = sum(rows,'ppto');
    const ingV = sum(rows,'ingreso_variable'); const pptV = sum(rows,'ppt_variable');
    const ingA = sum(rows,'ingreso_admin'); const ingF = sum(rows,'ingreso_fijo');
    const te = v>0 ? ((ingA+ingF+ingV)/v)*100 : null;
    const ac = areaActivaCalc(rows);
    const vxm2 = ac>0 ? v/ac : null;
    return {c, v, va, p, dv:delta(v,va), dp:pctOf(v,p), cumplV:pctOf(ingV,pptV), te, vxm2};
  });
  data.sort((a,b) => {
    const va = a[rSort.col] ?? -Infinity;
    const vb = b[rSort.col] ?? -Infinity;
    return (va < vb ? -1 : va > vb ? 1 : 0) * rSort.dir;
  });
  const cols = [
    {key:'c', label:'CECO'}, {key:'v', label:'Venta 2026'}, {key:'vxm2', label:'Venta/m²'},
    {key:'dv', label:'vs 2025'}, {key:'dp', label:'vs PPT Ventas'},
    {key:'cumplV', label:'Cumpl. Ing. Variable'}, {key:'te', label:'TE'}
  ];
  const thHtml = cols.map(col => {
    const cls = col.key !== 'c' ? ' class="num"' : '';
    const sortCls = rSort.col === col.key ? (rSort.dir > 0 ? ' sort-asc' : ' sort-desc') : '';
    return `<th${cls} class="${sortCls.trim()}" onclick="sortRanking('${col.key}')">${col.label}</th>`;
  }).join('');
  el.innerHTML = `<table class="rtbl"><thead><tr>${thHtml}</tr></thead><tbody>` +
  data.map(d => `<tr>
    <td class="ceco-name">${cecoShort(d.c)}</td>
    <td class="num">${fv(d.v)}</td>
    <td class="num" style="font-size:11px">${fm2(d.vxm2)}</td>
    <td class="num">${chip(d.dv)}</td>
    <td class="num">${chipPpto(d.dp)}</td>
    <td class="num">${chipPpto(d.cumplV)}</td>
    <td class="num">${chipTE(d.te)}</td>
  </tr>`).join('') + '</tbody></table>';
}

function renderCatPieG() {
  dc('ch-cat-pie');
  const rows = filtered();
  const cats = {};
  rows.forEach(r => { cats[r.categoria] = (cats[r.categoria]||0) + (r.venta||0); });
  const sorted = Object.entries(cats).sort((a,b)=>b[1]-a[1]);
  const ctx = document.getElementById('ch-cat-pie').getContext('2d');
  charts['ch-cat-pie'] = new Chart(ctx, {
    type:'doughnut',
    data:{labels:sorted.map(x=>x[0]), datasets:[{data:sorted.map(x=>x[1]),
      backgroundColor:['#0099e1','#00ad68','#ffd13f','#ff7c00','#b21e28','#7c3aed',
        '#0d9488','#f59e0b','#ec4899','#6b7280','#84cc16'],
      borderWidth:2, borderColor:'#fff'}]},
    options:{responsive:true, maintainAspectRatio:false, cutout:'60%',
      plugins:{legend:{position:'right',labels:{font:{family:'Sora',size:10},boxWidth:12,padding:8}},
        tooltip:{callbacks:{label:c=>' '+c.label+': '+fv(c.raw)+' ('+pctOf(c.raw,sum(rows,'venta')).toFixed(1)+'%)'}}}}
  });
}

function renderPptoG() {
  dc('ch-ppto-g');
  const data = D.cecos.map(c => {
    const rows = filtered(r=>r.ceco===c);
    const v = sum(rows,'venta'); const p = sum(rows,'ppto');
    return {c: cecoShort(c), pct: pctOf(v,p)||0};
  }).sort((a,b)=>b.pct-a.pct);
  const ctx = document.getElementById('ch-ppto-g').getContext('2d');
  charts['ch-ppto-g'] = new Chart(ctx, {
    type:'bar',
    data:{labels:data.map(d=>d.c), datasets:[{
      label:'Cumpl. PPT (%)',
      data:data.map(d=>d.pct),
      backgroundColor:data.map(d=>d.pct>=100?'#00ad6888':d.pct>=90?'#ffd13f88':'#b21e2888'),
      borderColor:data.map(d=>d.pct>=100?'#00ad68':d.pct>=90?'#ffd13f':'#b21e28'),
      borderWidth:1.5, borderRadius:4
    }]},
    options:{indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{legend:{display:false},
        tooltip:{callbacks:{label:c=>c.raw.toFixed(1)+'%'}}},
      scales:{x:{min:0, ticks:{callback:v=>v+'%', font:{family:'DM Mono',size:10}}, grid:{color:'#f0f0f0'}},
        y:{ticks:{font:{family:'Sora',size:10}}},
        ...({})}
    }
  });
}

function renderIngresosG() {
  dc('ch-ing-g');
  const data = D.cecos.map(c => {
    const rows = filtered(r=>r.ceco===c);
    return {c:cecoShort(c), admin:sum(rows,'ingreso_admin'), fijo:sum(rows,'ingreso_fijo'),
      variable:sum(rows,'ingreso_variable'), pptV:sum(rows,'ppt_variable')};
  });
  const ctx = document.getElementById('ch-ing-g').getContext('2d');
  charts['ch-ing-g'] = new Chart(ctx, {
    type:'bar',
    data:{labels:data.map(d=>d.c), datasets:[
      {label:'Admin', data:data.map(d=>d.admin), backgroundColor:'#0099e155', borderColor:CC[0], borderWidth:1.5, borderRadius:3, stack:'ing'},
      {label:'Fijo',  data:data.map(d=>d.fijo),  backgroundColor:'#00ad6855', borderColor:CC[2], borderWidth:1.5, borderRadius:3, stack:'ing'},
      {label:'Variable', data:data.map(d=>d.variable), backgroundColor:'#ffd13f88', borderColor:CC[1], borderWidth:1.5, borderRadius:3, stack:'ing'},
    ]},
    options:{responsive:true, maintainAspectRatio:false,
      plugins:{legend:{labels:{font:{family:'Sora',size:10},boxWidth:12}},
        tooltip:{mode:'index', intersect:false, callbacks:{label:c=>' '+c.dataset.label+': '+fv(c.raw)}}},
      scales:{y:{stacked:true, ticks:{callback:v=>fv(v), font:{family:'DM Mono',size:10}}, grid:{color:'#f0f0f0'}},
        x:{stacked:true, ticks:{font:{family:'Sora',size:10}}}}}
  });
}

function renderTopMarcasG() {
  const rows = filtered();
  const byMarca = groupBy(rows, r => r.ceco+'|'+r.marca);
  const withDelta = byMarca.map(r => ({...r, dv: delta(r.venta, r.venta_ant)}))
    .filter(r => r.venta_ant > 0 && r.dv !== null && r.marca_foco === 'SI');

  const top = [...withDelta].sort((a,b)=>b.dv-a.dv).slice(0,10);
  const bot = [...withDelta].sort((a,b)=>a.dv-b.dv).slice(0,10);

  const rowHtml = (r, pos) => `
    <div class="brow">
      <div class="brow-left">
        <span class="brow-name">${pos+1}. ${r.marca}</span>
        <span class="brow-sub">${cecoShort(r.ceco)} · ${r.categoria||''}</span>
      </div>
      <div class="brow-right">
        <span class="brow-val">${fv(r.venta)}</span>
        ${chip(r.dv)}
        ${r.marca_foco==='SI'?'<span class="badge-foco">FOCO</span>':''}
      </div>
    </div>`;

  document.getElementById('top-crec').innerHTML  = top.map(rowHtml).join('') || '<p style="color:var(--txt3);font-size:12px">Sin datos comparables</p>';
  document.getElementById('top-caida').innerHTML = bot.map(rowHtml).join('') || '<p style="color:var(--txt3);font-size:12px">Sin datos comparables</p>';
}

function renderComparativoG() {
  dc('ch-comp-g');
  const allRows = filteredAllMeses();
  const labels = D.meses.map(m => MES_SHORT[m]);
  const datasets = D.cecos.map((c,i) => ({
    label: cecoShort(c),
    data: D.meses.map(m => sum(allRows.filter(r=>r.ceco===c && r.mes===m), 'venta')),
    borderColor: CC[i%CC.length], backgroundColor: CC[i%CC.length]+'33',
    tension:.35, pointRadius:3, borderWidth:2, fill:false
  }));
  const ctx = document.getElementById('ch-comp-g').getContext('2d');
  charts['ch-comp-g'] = new Chart(ctx, {
    type:'line', data:{labels, datasets},
    options:{responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{labels:{font:{family:'Sora',size:10},boxWidth:14}},
        tooltip:{mode:'index', intersect:false, callbacks:{
          label: c => '  ' + c.dataset.label + ': ' + fv(c.raw)
        }}
      },
      scales:{y:{ticks:{callback:v=>fv(v), font:{family:'DM Mono',size:10}}, grid:{color:'#f0f0f0'}},
        x:{ticks:{font:{family:'Sora',size:11}}}}}
  });
}

function renderComparativoVM2G() {
  dc('ch-comp-vm2-g');
  const allRows = filteredAllMeses();
  const labels = D.meses.map(m => MES_SHORT[m]);
  const datasets = D.cecos.map((c,i) => ({
    label: cecoShort(c),
    data: D.meses.map(m => {
      const rows = allRows.filter(r => r.ceco===c && r.mes===m);
      const v = sum(rows, 'venta');
      const ac = areaActivaCalc(rows);
      return ac > 0 ? v / ac : null;
    }),
    borderColor: CC[i%CC.length], backgroundColor: CC[i%CC.length]+'33',
    tension:.35, pointRadius:3, borderWidth:2, fill:false, spanGaps:true
  }));
  const ctx = document.getElementById('ch-comp-vm2-g').getContext('2d');
  charts['ch-comp-vm2-g'] = new Chart(ctx, {
    type:'line', data:{labels, datasets},
    options:{responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{labels:{font:{family:'Sora',size:10},boxWidth:14}},
        tooltip:{mode:'index', intersect:false, callbacks:{
          label: c => c.raw !== null ? '  ' + c.dataset.label + ': ' + fm2(c.raw) : null
        }}
      },
      scales:{y:{ticks:{callback:v=>fv(v)+'/m²', font:{family:'DM Mono',size:10}}, grid:{color:'#f0f0f0'}},
        x:{ticks:{font:{family:'Sora',size:11}}}}}
  });
}

/* ════════════════════════════════════
   VISTA CECO
════════════════════════════════════ */
function renderCecoView() {
  document.getElementById('ceco-titulo').textContent = ceco;
  renderKpisVentasC();
  renderKpisIngresosC();
  renderKpisTraficoC();
  renderEvolC();
  renderIngEvol();
  renderCatC();
  renderIngC();
  renderTopMarcasC();
  renderMarcaTable();
}

function renderKpisVentasC() {
  const rows = filteredCeco(ceco);
  const v = sum(rows,'venta'); const va = sum(rows,'venta_ant'); const p = sum(rows,'ppto');
  const area     = areaActiva(rows);      // para mostrar
  const areaCalc = areaActivaCalc(rows);  // para calcular VM2
  const dvsAnt = delta(v,va); const dvsPpto = pctOf(v,p);
  const vxm2 = areaCalc>0 ? v/areaCalc : null;
  const pptoVm2 = calcPptoVm2(filteredPptoVm2Ceco(ceco));
  const marcasRep = new Set(rows.filter(r=>r.venta>0).map(r=>r.marca)).size;
  const marcasFoco = new Set(rows.filter(r=>r.venta>0&&r.marca_foco==='SI').map(r=>r.marca)).size;

  document.getElementById('kpis-ventas-c').innerHTML = `
    <div class="kpi">
      <div class="kpi-label">Venta ${mesLabel()}</div>
      <div class="kpi-val sm">${fv(v)}</div>
      <div class="kpi-sub">${fvFull(v)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">vs Año Anterior 2025</div>
      <div class="kpi-val sm">${chip(dvsAnt)}</div>
      <div class="kpi-sub">2025: ${fv(va)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">vs Presupuesto</div>
      <div class="kpi-val sm">${chipPpto(dvsPpto)}</div>
      <div class="kpi-sub">PPT: ${fv(p)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Venta / m²</div>
      <div class="kpi-val xs">${fm2(vxm2)}</div>
      <div class="kpi-sub">Área: ${fn(area)} m²</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Presupuesto VM2</div>
      <div class="kpi-val xs">${fm2(pptoVm2)}</div>
      <div class="kpi-sub">${chipPpto(pctOf(vxm2, pptoVm2))} cumpl.</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Marcas Reportando</div>
      <div class="kpi-val">${marcasRep}</div>
      <div class="kpi-sub">con venta &gt; 0</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Marcas Foco</div>
      <div class="kpi-val">${marcasFoco}</div>
      <div class="kpi-sub">con venta &gt; 0</div>
    </div>
  `;
}

function renderKpisIngresosC() {
  const rows = filteredCeco(ceco);
  const admin = sum(rows,'ingreso_admin'); const fijo = sum(rows,'ingreso_fijo');
  const varR  = sum(rows,'ingreso_variable'); const pptV = sum(rows,'ppt_variable');
  const area     = areaActiva(rows);      // para mostrar
  const areaCalc = areaActivaCalc(rows);  // para calcular /m²
  const v = sum(rows,'venta');
  const total = admin+fijo+varR; const te = v>0?(total/v)*100:null;
  const dvV = pctOf(varR,pptV);

  document.getElementById('kpis-ingresos-c').innerHTML = `
    <div class="kpi kpi-ing">
      <div class="kpi-label">Ingreso Total</div>
      <div class="kpi-val sm">${fv(total)}</div>
      <div class="kpi-sub">Admin+Fijo+Variable</div>
    </div>
    <div class="kpi kpi-te">
      <div class="kpi-label">Tasa de Esfuerzo (TE)</div>
      <div class="kpi-val sm">${chipTE(te)}</div>
      <div class="kpi-sub">Ing.Total / Venta 2026</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">Ingreso Admin</div>
      <div class="kpi-val sm">${fv(admin)}</div>
      <div class="kpi-sub">${fm2(areaCalc>0?admin/areaCalc:null)}</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">Ingreso Fijo</div>
      <div class="kpi-val sm">${fv(fijo)}</div>
      <div class="kpi-sub">${fm2(areaCalc>0?fijo/areaCalc:null)}</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">Ing. Variable Real</div>
      <div class="kpi-val sm">${fv(varR)}</div>
      <div class="kpi-sub">${fm2(areaCalc>0?varR/areaCalc:null)}</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">Cumpl. Variable</div>
      <div class="kpi-val sm">${chipPpto(dvV)}</div>
      <div class="kpi-sub">PPT: ${fv(pptV)}</div>
    </div>
    <div class="kpi kpi-ing">
      <div class="kpi-label">Ingreso Total / m²</div>
      <div class="kpi-val xs">${fm2(areaCalc>0?total/areaCalc:null)}</div>
      <div class="kpi-sub">Área: ${fn(area)} m²</div>
    </div>
  `;
}

function renderEvolC() {
  dc('ch-evol-c');
  const allRows = filteredAllMeses(r=>r.ceco===ceco);
  const labels = D.meses.map(m=>MES_SHORT[m]);
  const venta  = D.meses.map(m=>sum(allRows.filter(r=>r.mes===m),'venta'));
  const ventaA = D.meses.map(m=>sum(allRows.filter(r=>r.mes===m),'venta_ant'));
  const ppto   = D.meses.map(m=>sum(allRows.filter(r=>r.mes===m),'ppto'));
  const marcasXMes = D.meses.map(m =>
    new Set(allRows.filter(r=>r.mes===m && r.venta>0).map(r=>r.marca)).size
  );
  const varPctC = D.meses.map((_,i) => ventaA[i] ? (venta[i] - ventaA[i]) / ventaA[i] * 100 : null);
  const ctx = document.getElementById('ch-evol-c').getContext('2d');
  charts['ch-evol-c'] = new Chart(ctx, {
    type:'line', data:{labels, datasets:[
      {label:'Venta 2026', data:venta, borderColor:CC[0], backgroundColor:CC[0]+'22', tension:.35, fill:true, pointRadius:4, borderWidth:2},
      {label:'Venta 2025', data:ventaA, borderColor:CC[2], borderDash:[5,3], tension:.35, pointRadius:3, borderWidth:1.5},
      {label:'PPT', data:ppto, borderColor:CC[3], borderDash:[3,3], tension:.35, pointRadius:3, borderWidth:1.5},
    ]},
    options:{responsive:true, maintainAspectRatio:false,
      plugins:{legend:{labels:{font:{family:'Sora',size:11},boxWidth:14}},
        tooltip:{mode:'index', intersect:false, callbacks:{
          label: c => '  ' + c.dataset.label + ': ' + fvFull(c.raw),
          footer: items => {
            if (!items.length) return '';
            const idx = items[0].dataIndex;
            const v = varPctC[idx];
            const pctStr = v !== null ? (v >= 0 ? '+' : '') + v.toFixed(1) + '% 2026 vs 2025' : '';
            return ['Marcas reportando: ' + marcasXMes[idx], pctStr].filter(Boolean).join('\n');
          }
        }}},
      scales:{y:{ticks:{callback:v=>fv(v), font:{family:'DM Mono',size:10}}, grid:{color:'#f0f0f0'}},
        x:{ticks:{font:{family:'Sora',size:11}}}}}
  });
}

function renderIngEvol() {
  dc('ch-ing-evol-c');
  const allRows = filteredAllMeses(r=>r.ceco===ceco);
  const labels = D.meses.map(m=>MES_SHORT[m]);
  const admin   = D.meses.map(m=>sum(allRows.filter(r=>r.mes===m),'ingreso_admin'));
  const fijo    = D.meses.map(m=>sum(allRows.filter(r=>r.mes===m),'ingreso_fijo'));
  const varR    = D.meses.map(m=>sum(allRows.filter(r=>r.mes===m),'ingreso_variable'));
  const ctx = document.getElementById('ch-ing-evol-c').getContext('2d');
  charts['ch-ing-evol-c'] = new Chart(ctx, {
    type:'bar', data:{labels, datasets:[
      {label:'Admin', data:admin, backgroundColor:'#0099e155', borderColor:CC[0], borderWidth:1.5, borderRadius:3, stack:'ing'},
      {label:'Fijo',  data:fijo,  backgroundColor:'#00ad6855', borderColor:CC[2], borderWidth:1.5, borderRadius:3, stack:'ing'},
      {label:'Variable', data:varR, backgroundColor:'#ffd13f88', borderColor:CC[1], borderWidth:1.5, borderRadius:3, stack:'ing'},
    ]},
    options:{responsive:true, maintainAspectRatio:false,
      plugins:{legend:{labels:{font:{family:'Sora',size:10},boxWidth:12}},
        tooltip:{mode:'index', intersect:false, callbacks:{label:c=>' '+c.dataset.label+': '+fv(c.raw)}}},
      scales:{y:{stacked:true, ticks:{callback:v=>fv(v), font:{family:'DM Mono',size:10}}, grid:{color:'#f0f0f0'}},
        x:{stacked:true, ticks:{font:{family:'Sora',size:11}}}}}
  });
}

function renderCatC() {
  dc('ch-cat-c');
  const rows = filteredCeco(ceco);
  const cats = {};
  rows.forEach(r => { cats[r.categoria]=(cats[r.categoria]||{v:0,va:0,p:0});
    cats[r.categoria].v+=r.venta||0; cats[r.categoria].va+=r.venta_ant||0; cats[r.categoria].p+=r.ppto||0; });
  const sorted = Object.entries(cats).sort((a,b)=>b[1].v-a[1].v);
  const labels = sorted.map(x=>x[0]);
  const ctx = document.getElementById('ch-cat-c').getContext('2d');
  charts['ch-cat-c'] = new Chart(ctx, {
    type:'bar', data:{labels, datasets:[
      {label:'Venta 2026', data:sorted.map(x=>x[1].v), backgroundColor:'#0099e166', borderColor:CC[0], borderWidth:1.5, borderRadius:3},
      {label:'Venta 2025', data:sorted.map(x=>x[1].va), backgroundColor:'#00ad6844', borderColor:CC[2], borderWidth:1.5, borderRadius:3},
      {label:'PPT', data:sorted.map(x=>x[1].p), type:'line', borderColor:CC[3], pointRadius:4, borderWidth:2, fill:false},
    ]},
    options:{responsive:true, maintainAspectRatio:false,
      plugins:{legend:{labels:{font:{family:'Sora',size:10},boxWidth:12}},
        tooltip:{callbacks:{label:c=>' '+c.dataset.label+': '+fv(c.raw)}}},
      scales:{y:{ticks:{callback:v=>fv(v), font:{family:'DM Mono',size:10}}, grid:{color:'#f0f0f0'}},
        x:{ticks:{font:{family:'Sora',size:10},maxRotation:30}}}}
  });
}

function renderIngC() {
  dc('ch-ing-c');
  const rows = filteredCeco(ceco);
  const admin = sum(rows,'ingreso_admin'); const fijo = sum(rows,'ingreso_fijo');
  const varR = sum(rows,'ingreso_variable'); const pptV = sum(rows,'ppt_variable');
  const ctx = document.getElementById('ch-ing-c').getContext('2d');
  charts['ch-ing-c'] = new Chart(ctx, {
    type:'bar',
    data:{labels:['Admin','Fijo','Variable Real','PPT Variable'],
      datasets:[{data:[admin,fijo,varR,pptV],
        backgroundColor:['#0099e188','#00ad6888','#ffd13faa','#ff7c0044'],
        borderColor:[CC[0],CC[2],CC[1],CC[3]], borderWidth:2, borderRadius:6}]},
    options:{responsive:true, maintainAspectRatio:false,
      plugins:{legend:{display:false},
        tooltip:{callbacks:{label:c=>fv(c.raw)}}},
      scales:{y:{ticks:{callback:v=>fv(v), font:{family:'DM Mono',size:10}}, grid:{color:'#f0f0f0'}},
        x:{ticks:{font:{family:'Sora',size:11}}}}}
  });
}

function renderTopMarcasC() {
  const rows = filteredCeco(ceco);
  const byMarca = groupBy(rows, r=>r.marca);
  const withDelta = byMarca.map(r => ({...r, dv: delta(r.venta, r.venta_ant)}))
    .filter(r => r.venta_ant > 0 && r.dv !== null && r.marca_foco === 'SI');

  const top = [...withDelta].sort((a,b)=>b.dv-a.dv).slice(0,10);
  const bot = [...withDelta].sort((a,b)=>a.dv-b.dv).slice(0,10);

  const rowHtml = (r, pos) => `
    <div class="brow">
      <div class="brow-left">
        <span class="brow-name">${pos+1}. ${r.marca}</span>
        <span class="brow-sub">${r.categoria||''}</span>
      </div>
      <div class="brow-right">
        <span class="brow-val">${fv(r.venta)}</span>
        ${chip(r.dv)}
        <span class="badge-foco">FOCO</span>
      </div>
    </div>`;

  document.getElementById('top-marcas-c').innerHTML = top.map(rowHtml).join('') || '<p style="color:var(--txt3);font-size:12px">Sin datos comparables</p>';
  document.getElementById('top-riesgo-c').innerHTML = bot.map(rowHtml).join('') || '<p style="color:var(--txt3);font-size:12px">Sin datos comparables</p>';
}

function sortMarcas(col) {
  if (mSort.col === col) mSort.dir *= -1;
  else { mSort.col = col; mSort.dir = -1; }
  document.querySelectorAll('.mtbl th').forEach(th => { th.classList.remove('sort-asc','sort-desc'); });
  renderMarcaTable();
}

function renderMarcaTable() {
  const rows = filteredCeco(ceco);
  const byMarca = groupBy(rows, r=>r.marca);

  byMarca.forEach(r => {
    // VM2 usa área acumulada (coherente con venta acumulada en modo comparar)
    const areaCalcMarca = areaActivaCalc(rows.filter(rr => rr.marca === r.marca));
    r.vxm2 = areaCalcMarca>0 ? r.venta/areaCalcMarca : null;
    r.dvs25 = delta(r.venta, r.venta_ant);
    r.dppto = pctOf(r.venta, r.ppto);
    r.dppt_var = pctOf(r.ingreso_variable, r.ppt_variable);
    r.te = r.venta>0 ? ((r.ingreso_admin+r.ingreso_fijo+r.ingreso_variable)/r.venta)*100 : null;
  });

  let filtered2 = byMarca;
  if (mSearch) filtered2 = filtered2.filter(r => r.marca.toLowerCase().includes(mSearch));

  filtered2.sort((a,b) => {
    const va = a[mSort.col]??-Infinity; const vb = b[mSort.col]??-Infinity;
    return (va<vb?-1:va>vb?1:0) * mSort.dir;
  });

  // update sort headers
  document.querySelectorAll('.mtbl th').forEach(th => {
    const col = th.getAttribute('onclick')?.match(/'(.+)'/)?.[1];
    if (col === mSort.col) th.classList.add(mSort.dir>0?'sort-asc':'sort-desc');
    else th.classList.remove('sort-asc','sort-desc');
  });

  const tbody = document.getElementById('mtbl-body');
  tbody.innerHTML = filtered2.map(r => `<tr>
    <td class="marca-cell" title="${r.marca}">${r.marca}</td>
    <td style="font-size:11px">${r.categoria||''}</td>
    <td class="num">${fn(r.area)}</td>
    <td class="num">${fvFull(r.venta)}</td>
    <td class="num">${fm2(r.vxm2)}</td>
    <td class="num">${fvFull(r.venta_ant)}</td>
    <td class="num">${chip(r.dvs25)}</td>
    <td class="num">${fvFull(r.ppto)}</td>
    <td class="num">${chipPpto(r.dppto)}</td>
    <td class="num">${fvFull(r.ingreso_admin)}</td>
    <td class="num">${fvFull(r.ingreso_fijo)}</td>
    <td class="num">${fvFull(r.ingreso_variable)}</td>
    <td class="num">${fvFull(r.ppt_variable)}</td>
    <td class="num">${chipPpto(r.dppt_var)}</td>
    <td class="num">${chipTE(r.te)}</td>
    <td style="text-align:center">${r.marca_foco==='SI'?'<span class="badge-foco">FOCO</span>':''}</td>
  </tr>`).join('');
}

// Llamar init() directamente (el script está al final del body, DOM ya listo)
try {
  init();
} catch(e) {
  document.body.insertAdjacentHTML('afterbegin',
    '<div style="background:#fee2e2;color:#991b1b;padding:16px;font-family:monospace;font-size:13px;z-index:9999;position:fixed;top:0;left:0;right:0;border-bottom:2px solid #b91c1c">'
    + '<b>Error JavaScript:</b> ' + e.message + '<br>' + (e.stack||'') + '</div>');
}
</script>
</body>
</html>"""


if __name__ == '__main__':
    main()
