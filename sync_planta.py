#!/usr/bin/env python3
"""
sync_planta.py — Lee Balance_Semanal.xlsx y actualiza la sección PLANTA de index.html
Ejecutar: python3 sync_planta.py  (desde la carpeta GESTION COSTOS)
"""

import openpyxl, re
from pathlib import Path
from collections import defaultdict

EXCEL = Path(__file__).parent / 'Balance_Semanal.xlsx'
HTML  = Path(__file__).parent / 'index.html'

SEM_DATES  = ['12/07','13/07','14/07','15/07','16/07','17/07','18/07']  # Semana 29
SEM30_DATES= ['19/07','20/07','21/07','22/07','23/07','24/07','25/07']  # Semana 30
ALL_DATES  = SEM_DATES + SEM30_DATES

# ── 1. Leer Excel ────────────────────────────────────────────────
# NOTA: se leen solo columnas de ENTRADA MANUAL (sacos, fecha, jabas de selección/
# despacho, chancho, ticket). Las columnas calculadas (kg, totales, saldo corrido,
# merma) se RECALCULAN aquí mismo en Python en vez de leer los valores cacheados
# de las fórmulas de Excel — así el script no depende de que el archivo haya sido
# abierto/recalculado en Excel/Numbers justo antes de correr sync_planta.py.
wb = openpyxl.load_workbook(EXCEL, data_only=True)
ws = wb.active

def to_fecha(v):
    """Convierte datetime o string a 'DD/MM'"""
    if v is None: return None
    if hasattr(v, 'day'): return f"{v.day:02d}/{v.month:02d}"
    if isinstance(v, str) and '/' in v and v[:2].isdigit(): return v[:5]
    return None

rows_raw = list(ws.iter_rows(min_row=5, values_only=True))

opening = {}
pedidos = []

# Estado corrido (saldo en planta + saldo chancho), se va acumulando fila a fila
run_A = run_P = run_C = run_M = 0
run_chSaldo = 0

for row in rows_raw:
    fecha_col  = to_fecha(row[0])   # col A: fecha cosecha
    dia_col    = str(row[1]) if row[1] else None
    fecha_desp = to_fecha(row[9])   # col J: fecha despacho
    guia       = row[10]

    # Detectar fila SEMANA (saltar, no parar — puede haber más semanas abajo)
    if isinstance(row[0], str) and 'SEMANA' in row[0].upper():
        continue

    # Fila apertura: día anterior al inicio de semana, con saldo pero sin ingreso numérico
    if dia_col and dia_col.lower() in ('sab','sáb') and not isinstance(row[2], (int, float)) and (row[17] is not None or row[28] is not None):
        run_A, run_P, run_C, run_M = row[17] or 0, row[18] or 0, row[19] or 0, row[20] or 0
        run_chSaldo = row[28] or 0
        opening = {
            'fecha': fecha_col,
            'saldoA': run_A, 'saldoP': run_P, 'saldoC': run_C, 'saldoM': run_M,
            'chSaldo': run_chSaldo,
        }
        continue

    # Determinar la fecha del registro
    # Prioridad: columna A; si es None, usar columna J (despacho-only rows)
    fecha = fecha_col or fecha_desp
    if fecha is None:
        continue

    # Ignorar filas fuera de la semana (ej: dom19 siguiente)
    if fecha not in ALL_DATES:
        continue

    dia = dia_col or ''

    ingr_s   = row[2]  or 0
    ingr_kg  = ingr_s * 95                                    # D = C*95
    prodA    = row[4]  or 0
    prodP    = row[5]  or 0
    prodC    = row[6]  or 0
    prodM    = row[7]  or 0
    prodTot  = prodA+prodP+prodC+prodM                        # I
    despA    = row[11] or 0
    despP    = row[12] or 0
    despC    = row[13] or 0
    despM    = row[14] or 0
    despTotJ = despA+despP+despC+despM                        # P
    despTotKg= despA*15+despP*18+despC*17+despM*17            # Q
    tras     = row[23] or 0
    chProd   = row[22] or 0
    chVenta  = row[25] or 0
    chTicket = str(row[27]) if row[27] else ''

    # Saldo en planta (jabas) — corrido desde la fila anterior
    run_A = run_A + prodA - despA - tras
    run_P = run_P + prodP - despP
    run_C = run_C + prodC - despC
    run_M = run_M + prodM - despM
    saldoA, saldoP, saldoC, saldoM = run_A, run_P, run_C, run_M
    saldoTot = saldoA+saldoP+saldoC+saldoM                     # V

    # Saldo chancho (sacos) — corrido
    run_chSaldo = run_chSaldo + chProd - chVenta
    chSaldo = run_chSaldo

    chProd_kg = chProd*95                                      # Y

    # dash = fila sin selección NI despacho aún: puede ser "solo despacho" (ingr_s=0)
    # o sacos comprados que todavía no se procesan (ingr_s>0 pero prodTot=0 y despTotJ=0).
    # En ambos casos no hay nada real que mostrar en Selección/Merma todavía.
    desp_only = (prodTot == 0 and despTotJ == 0)

    if ingr_s == 0 or desp_only:
        mermaKg = 0
        mermaPct = 0
    else:
        mermaKg = ingr_kg - (prodA*15+prodP*18+prodC*17+prodM*17) - chProd_kg
        mermaPct = (mermaKg/ingr_kg*100) if ingr_kg else 0

    pedidos.append({
        'fecha': fecha, 'dia': dia,
        'ingr_s': ingr_s, 'ingr_kg': ingr_kg,
        'prodA': prodA, 'prodP': prodP, 'prodC': prodC, 'prodM': prodM, 'prodTot': prodTot,
        'guia': str(guia) if guia else '—',
        'despA': despA, 'despP': despP, 'despC': despC, 'despM': despM,
        'despTotJ': despTotJ, 'despTotKg': int(despTotKg),
        'saldoA': saldoA, 'saldoP': saldoP, 'saldoC': saldoC, 'saldoM': saldoM,
        'saldoTot': saldoTot,
        'chProd': chProd, 'chVenta': chVenta, 'chTicket': chTicket,
        'chSaldo': chSaldo, 'tras': tras,
        'mermaKg': mermaKg,
        'mermaPct': mermaPct,
        'desp_only': desp_only,
    })

print(f"✅ Apertura: {opening}")
print(f"✅ Pedidos totales leídos: {len(pedidos)}")
for p in pedidos:
    print(f"   {p['fecha']} {p['dia']:4s}  ingr={p['ingr_s']}s  prodTot={p['prodTot']}j  chProd={p['chProd']}s  chSaldo={p['chSaldo']}s  tras={p['tras']}")

# Separar Semana 29 (para balance/sd/chancho) del resto (Sem30 solo entra en _md[])
sem29 = [p for p in pedidos if p['fecha'] in SEM_DATES]

# ── 2. Totales semanales (Semana 29 únicamente) ──────────────────
by_date = defaultdict(list)
for p in pedidos:  # by_date incluye TODAS las semanas (para _md[])
    by_date[p['fecha']].append(p)

def day_closing(fecha):
    ps = by_date.get(fecha, [])
    return ps[-1] if ps else None

sab18 = day_closing('18/07')
final_ch_saldo = sab18['chSaldo'] if sab18 else 0
final_saldo_tot= sab18['saldoTot'] if sab18 else 3
final_saldoA   = sab18['saldoA']   if sab18 else 3

ch_opening = opening.get('chSaldo', 0)
total_chProd  = sum(p['chProd']    for p in sem29)
total_chVenta = sum(p['chVenta']   for p in sem29)
total_ingr_s  = sum(p['ingr_s']   for p in sem29)
total_ingr_kg = sum(p['ingr_kg']  for p in sem29)
total_prodA   = sum(p['prodA']    for p in sem29)
total_prodP   = sum(p['prodP']    for p in sem29)
total_prodC   = sum(p['prodC']    for p in sem29)
total_prodM   = sum(p['prodM']    for p in sem29)
total_prodTot = sum(p['prodTot']  for p in sem29)
total_despA   = sum(p['despA']    for p in sem29)
total_despP   = sum(p['despP']    for p in sem29)
total_despC   = sum(p['despC']    for p in sem29)
total_despM   = sum(p['despM']    for p in sem29)
total_despTotJ= sum(p['despTotJ'] for p in sem29)
total_despTotKg=sum(p['despTotKg']for p in sem29)
total_mermaKg = sum(p['mermaKg'] for p in sem29 if not p['desp_only'])
total_merka_pct = (total_mermaKg / total_ingr_kg * 100) if total_ingr_kg > 0 else 0
new_ch_total = ch_opening + total_chProd

ch_tickets_str = ' + '.join(
    f"Tkt {p['chTicket']} ({p['chVenta']}s)"
    for p in sem29 if p['chVenta'] > 0 and p['chTicket']
)

print(f"\n✅ Totales:")
print(f"   Ingreso: {total_ingr_s}s / {total_ingr_kg:,} kg")
print(f"   Selección: {total_prodTot}j")
print(f"   Despacho: {total_despTotJ}j / {total_despTotKg:,} kg")
print(f"   Saldo cierre: A={final_saldoA}j, tot={final_saldo_tot}j")
print(f"   Chancho: aper={ch_opening}s + prod={total_chProd}s = {new_ch_total}s · venta={total_chVenta}s · saldo={final_ch_saldo}s")
print(f"   Merma: {total_mermaKg:+.0f} kg / {total_merka_pct:+.1f}%")

# ── 3. Generar HTML ──────────────────────────────────────────────
def ch_cell(prod, venta, ticket, saldo, bdr, ingr_kg=0):
    pct_txt = ''
    if prod > 0 and ingr_kg > 0:
        pct = prod*95/ingr_kg*100
        pct_txt = f'<div style="font-size:.6rem;font-weight:700;color:#c62828">{pct:.0f}% del ingreso</div>'
    P = (f'<td style="padding:5px 6px;text-align:center;color:#795548;border-bottom:{bdr};background:#fff8f0;line-height:1.4">'
         f'<div style="font-weight:700">{prod} s</div><div style="font-size:.63rem;color:#a1887f">{prod*95:,} kg</div>{pct_txt}</td>'
         if prod > 0 else
         f'<td style="padding:5px 6px;text-align:center;color:#bbb;border-bottom:{bdr};background:#fff8f0">—</td>')
    V = (f'<td style="padding:5px 6px;text-align:center;font-weight:800;color:#e65100;border-bottom:{bdr};background:#fff8f0;line-height:1.4">'
         f'<div>{venta} s</div><div style="font-size:.63rem;color:#e57373">{venta*95:,} kg</div>'
         f'<div style="font-size:.6rem;font-weight:600;color:#e57373">{ticket}</div></td>'
         if venta > 0 else
         f'<td style="padding:5px 6px;text-align:center;color:#bbb;border-bottom:{bdr};background:#fff8f0">—</td>')
    sc = '#4e342e' if saldo > 0 else '#bbb'
    S = (f'<td style="padding:5px 6px;text-align:center;font-weight:800;color:{sc};border-bottom:{bdr};background:#efebe9;line-height:1.4">'
         f'<div>{saldo} s</div><div style="font-size:.63rem;color:#a1887f">{saldo*95:,} kg</div></td>')
    return P+V+S

def ingr_cell(s, kg, bdr):
    if s == 0:
        return f'<td style="padding:5px 6px;text-align:center;border-bottom:{bdr};background:#e3f2fd;color:#bbb;font-size:.72rem">—</td>'
    return (f'<td style="padding:5px 6px;text-align:center;border-bottom:{bdr};background:#e3f2fd;line-height:1.5">'
            f'<div style="font-size:.73rem;font-weight:800;color:#1565c0">{s} s</div>'
            f'<div style="font-size:.64rem;font-weight:600;color:#1976d2">{kg:,} kg</div></td>')

def sel_cells(a,p,c,m,tot,bdr,dash=False):
    if dash:
        d='<td style="padding:7px 8px;text-align:center;color:#bbb;border-bottom:{b};background:#f3f8ff">—</td>'.format(b=bdr)
        return d*4 + f'<td style="padding:7px 8px;text-align:center;font-weight:800;color:#bbb;border-bottom:{bdr};background:#dbeafe">— j</td>'
    return (
        f'<td style="padding:7px 8px;text-align:center;color:#1565c0;border-bottom:{bdr};background:#f3f8ff">{a}</td>'
        f'<td style="padding:7px 8px;text-align:center;color:#1565c0;border-bottom:{bdr};background:#f3f8ff">{p}</td>'
        f'<td style="padding:7px 8px;text-align:center;color:#1565c0;border-bottom:{bdr};background:#f3f8ff">{c}</td>'
        f'<td style="padding:7px 8px;text-align:center;color:#1565c0;border-bottom:{bdr};background:#f3f8ff">{m}</td>'
        f'<td style="padding:7px 8px;text-align:center;font-weight:800;color:#0d47a1;border-bottom:{bdr};background:#dbeafe">{tot}</td>'
    )

def desp_cells(guia,a,p,c,m,totj,totkg,bdr):
    return (
        f'<td style="padding:5px 6px;text-align:center;color:#555;font-size:.7rem;border-bottom:{bdr};font-family:monospace;background:#fff9f0;font-weight:700">{guia}</td>'
        f'<td style="padding:7px 8px;text-align:center;color:#2e7d32;border-bottom:{bdr};background:#f3faf3">{a}</td>'
        f'<td style="padding:7px 8px;text-align:center;color:#2e7d32;border-bottom:{bdr};background:#f3faf3">{p}</td>'
        f'<td style="padding:7px 8px;text-align:center;color:#2e7d32;border-bottom:{bdr};background:#f3faf3">{c}</td>'
        f'<td style="padding:7px 8px;text-align:center;color:#2e7d32;border-bottom:{bdr};background:#f3faf3">{m}</td>'
        f'<td style="padding:7px 8px;text-align:center;font-weight:800;color:#1b5e20;border-bottom:{bdr};background:#dcfce7">{totj} j</td>'
        f'<td style="padding:7px 8px;text-align:center;font-weight:700;color:#1b5e20;border-bottom:{bdr};background:#dcfce7">{totkg:,} kg</td>'
    )

def saldo_cells(a,p,c,m,tot,bdr):
    def cl(v, bold=False):
        col='#4527a0' if v>0 else '#aaa'
        fw='font-weight:900;' if bold else ''
        bg='#ede7f6' if bold else '#f5f0ff'
        return f'<td style="padding:7px 8px;text-align:center;{fw}color:{col};border-bottom:{bdr};background:{bg}">{v if v>0 else 0}</td>'
    return cl(a)+cl(p)+cl(c)+cl(m)+f'<td style="padding:7px 8px;text-align:center;font-weight:900;color:#311b92;border-bottom:{bdr};background:#ede7f6">{tot} j</td>'

def merma_cells(kg, pct, bdr):
    if kg==0 and pct==0:
        return (f'<td style="padding:7px 8px;text-align:center;color:#bbb;border-bottom:{bdr};background:#fff3e0">—</td>'
                f'<td style="padding:7px 8px;text-align:center;color:#bbb;border-bottom:{bdr};background:#fff3e0">—</td>')
    color='#1b5e20' if kg>=0 else '#e65100'
    sign='+' if kg>=0 else '-'
    pct_val = abs(float(pct)) if float(pct) != 0 else 0.0
    return (f'<td style="padding:7px 8px;text-align:center;font-weight:800;color:{color};border-bottom:{bdr};background:#fff3e0">{int(round(kg)):+d} kg</td>'
            f'<td style="padding:7px 8px;text-align:center;font-weight:800;color:{color};border-bottom:{bdr};background:#fff3e0">{sign}{pct_val:.1f}%</td>')

def make_balance_rows():
    out = ''
    for i, p in enumerate(sem29):
        f, dia = p['fecha'], p['dia']
        is_second = (i > 0 and sem29[i-1]['fecha'] == f)
        
        if f == '12/07':
            bg='#f3e5f5'; bdr='1px solid #e1bee7'
            dcol='#6a1b9a'; diacol='#7b1fa2'; diafw='font-weight:700;'
        elif f == '18/07':
            bg='#e3f2fd'; bdr='1px solid #bbdefb'
            dcol='#1565c0'; diacol='#1565c0'; diafw='font-weight:700;'
        elif is_second:
            bg='#f0f8ff'; bdr='1px solid #eee'
            dcol='#1565c0'; diacol='#888'; diafw=''
        elif i%2==1: bg='#fafafa'; bdr='1px solid #eee'; dcol='#333'; diacol='#888'; diafw=''
        else:        bg='#fff';    bdr='1px solid #eee'; dcol='#333'; diacol='#888'; diafw=''

        row = f'        <tr style="background:{bg}">\n'
        if is_second:
            row += f'          <td style="padding:7px 10px;text-align:center;font-size:.72rem;color:#1565c0;border-bottom:{bdr};font-style:italic">↳ Ped.2</td>\n'
            row += f'          <td style="padding:7px 6px;text-align:center;color:{diacol};font-size:.75rem;{diafw}border-bottom:{bdr}">{dia}</td>'
        else:
            row += f'          <td style="padding:7px 10px;text-align:center;font-weight:800;font-size:.82rem;color:{dcol};border-bottom:{bdr}">{f}</td>\n'
            row += f'          <td style="padding:7px 6px;text-align:center;color:{diacol};font-size:.75rem;{diafw}border-bottom:{bdr}">{dia}</td>'

        row += ingr_cell(p['ingr_s'], p['ingr_kg'], bdr)
        row += sel_cells(p['prodA'],p['prodP'],p['prodC'],p['prodM'],p['prodTot'],bdr,dash=p['desp_only'])
        row += desp_cells(p['guia'],p['despA'],p['despP'],p['despC'],p['despM'],p['despTotJ'],p['despTotKg'],bdr)
        row += saldo_cells(p['saldoA'],p['saldoP'],p['saldoC'],p['saldoM'],p['saldoTot'],bdr)
        row += ch_cell(p['chProd'],p['chVenta'],p['chTicket'],p['chSaldo'],bdr,ingr_kg=p['ingr_kg'])
        if p['desp_only'] or p['ingr_s']==0:
            row += merma_cells(0,0,bdr)
        else:
            row += merma_cells(p['mermaKg'],p['mermaPct'],bdr)
        row += '\n        </tr>\n'

        # Fila de trasiego si aplica
        if p['tras'] and p['tras'] > 0:
            row += (f'        <tr style="background:#fff8e1">'
                    f'<td colspan="3" style="padding:4px 10px;font-size:.68rem;color:#e65100;border-bottom:{bdr};font-weight:700">↳ TRASIEGO</td>'
                    f'<td colspan="5" style="padding:4px 8px;font-size:.68rem;color:#795548;border-bottom:{bdr};font-style:italic">'
                    f'−{p["tras"]} jaba(s) trasigeada(s) durante llenado</td>'
                    f'<td colspan="7" style="padding:4px 8px;font-size:.68rem;color:#555;border-bottom:{bdr}">Stock ajustado calidad A: −{p["tras"]} j (−{p["tras"]*15} kg)</td>'
                    f'<td colspan="5" style="border-bottom:{bdr}"></td>'
                    f'<td colspan="3" style="border-bottom:{bdr}"></td>'
                    f'<td colspan="2" style="border-bottom:{bdr}"></td>'
                    f'</tr>\n')
        out += row
    return out

balance_rows = make_balance_rows()

# SEMANA row
ch_sc = '#2e7d32' if final_ch_saldo==0 else '#4e342e'
ch_sb = '#c8e6c9' if final_ch_saldo==0 else '#efebe9'
mc = '#1b5e20' if total_mermaKg>=0 else '#e65100'
ms = '+' if total_mermaKg>=0 else '-'
semana_row = f'''        <tr style="background:#eceff1;border-top:2px solid #90a4ae">
          <td colspan="3" style="padding:6px 10px;text-align:center;font-weight:900;font-size:.85rem;color:#263238;line-height:1.5">SEMANA 29<br><span style="font-size:.67rem;font-weight:600;color:#1565c0">{total_ingr_s} s · {total_ingr_kg:,} kg ingresados</span></td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#0d47a1;background:#dbeafe">{total_prodA}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#0d47a1;background:#dbeafe">{total_prodP}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#0d47a1;background:#dbeafe">{total_prodC}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#0d47a1;background:#dbeafe">{total_prodM}</td>
          <td style="padding:8px;text-align:center;font-weight:900;color:#0d47a1;background:#bfdbfe">{total_prodTot:,} j</td>
          <td style="padding:8px;text-align:center;color:#bbb;font-size:.7rem;background:#fff9f0">—</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#1b5e20;background:#dcfce7">{total_despA}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#1b5e20;background:#dcfce7">{total_despP}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#1b5e20;background:#dcfce7">{total_despC}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#1b5e20;background:#dcfce7">{total_despM}</td>
          <td style="padding:8px;text-align:center;font-weight:900;color:#1b5e20;background:#bbf7d0">{total_despTotJ:,} j</td>
          <td style="padding:8px;text-align:center;font-weight:900;color:#1b5e20;background:#bbf7d0">{total_despTotKg:,} kg</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#311b92;background:#ede7f6">{final_saldoA}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#aaa;background:#ede7f6">0</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#aaa;background:#ede7f6">0</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#aaa;background:#ede7f6">0</td>
          <td style="padding:8px;text-align:center;font-weight:900;color:#311b92;background:#ddd6fe">{final_saldo_tot} j</td>
          <td style="padding:6px 4px;text-align:center;font-weight:800;color:#4e342e;background:#efebe9;line-height:1.4"><div>{total_chProd} s prod.</div><div style="font-size:.66rem;color:#a1887f">{total_chProd*95:,} kg</div><div style="font-size:.6rem;font-weight:700;color:#c62828">{(total_chProd*95/total_ingr_kg*100) if total_ingr_kg>0 else 0:.1f}% del ingreso</div></td><td style="padding:6px 4px;text-align:center;font-weight:900;color:#e65100;background:#fff8f0;line-height:1.4"><div>{total_chVenta} s</div><div style="font-size:.66rem;color:#e57373">{total_chVenta*95:,} kg</div></td><td style="padding:6px 4px;text-align:center;font-weight:900;color:{ch_sc};background:{ch_sb};line-height:1.4"><div>{final_ch_saldo} s</div><div style="font-size:.66rem;color:#a5d6a7">{final_ch_saldo*95:,} kg</div></td>
          <td style="padding:8px;text-align:center;font-weight:900;color:{mc};background:#fff3e0">{int(round(total_mermaKg)):+d} kg</td>
          <td style="padding:8px;text-align:center;font-weight:900;color:{mc};background:#fff3e0">{ms}{abs(total_merka_pct):.1f}%</td>
        </tr>'''

# ── 4. _md[] array ───────────────────────────────────────────────
def prod_kg(a,p,c,m): return a*15+p*18+c*17+m*17
MD_ORDER = ['13/07','14/07','15/07','16/07','17/07','18/07','12/07','19/07']
# Extiende automáticamente con cualquier fecha de Semana 30 (después del 19/07)
# que ya tenga datos cargados, para que el panel "Resumen del día" del DIARIO
# funcione también en esos días sin tener que tocar este código cada vez.
MD_ORDER += [d for d in SEM30_DATES[1:] if by_date.get(d) and d not in MD_ORDER]
md_entries = []
for fecha in MD_ORDER:
    ps = by_date.get(fecha,[])
    if not ps or all(p['ingr_s']==0 for p in ps):
        # Despacho only
        if ps:
            dp = sum(p['despTotKg'] for p in ps)
        else:
            dp = 0
        md_entries.append(f"  {{f:'{fecha}',s:0,ing:0,j:0,jn:0,ch_s:0,ch:0,sal:0,mer:0,pct:0,desp:{dp}}}")
        continue
    ingr_s  = sum(p['ingr_s']  for p in ps)
    ingr_kg = sum(p['ingr_kg'] for p in ps)
    pA=sum(p['prodA'] for p in ps); pP=sum(p['prodP'] for p in ps)
    pC=sum(p['prodC'] for p in ps); pM=sum(p['prodM'] for p in ps)
    jTot=sum(p['prodTot'] for p in ps)
    jkg=prod_kg(pA,pP,pC,pM)
    chs=sum(p['chProd'] for p in ps); chkg=chs*95
    sal=jkg+chkg; mer=ingr_kg-sal
    pct=round(mer/ingr_kg*100,1) if ingr_kg>0 else 0
    dp=sum(p['despTotKg'] for p in ps if p['desp_only'])
    md_entries.append(f"  {{f:'{fecha}',s:{ingr_s},ing:{ingr_kg},j:{jkg},jn:{jTot},ch_s:{chs},ch:{chkg},sal:{sal},mer:{mer},pct:{pct},desp:{int(dp)}}}")

new_md_block = "var md=window._md||[\n"+",\n".join(md_entries)+"\n];"

# ── 5. Chancho alert ─────────────────────────────────────────────
ch_alert = (
    f'<div style="display:flex;align-items:flex-start;gap:8px;background:#e8f5e9;border:1.5px solid #43a047;border-radius:8px;padding:10px 14px;margin-top:10px">\n'
    f'      <span>✅</span>\n'
    f'      <div>\n'
    f'        <div style="font-weight:800;color:#1b5e20;font-size:.75rem">CHANCHO AL DÍA · SEM. 29 · TODOS CON TICKET</div>\n'
    f'        <div style="font-size:.68rem;color:#388e3c;margin-top:3px">Apertura: {ch_opening} s · Producidos sem: {total_chProd} s · Total: {new_ch_total} s · {ch_tickets_str}</div>\n'
    f'        <div style="font-size:.67rem;color:#555;margin-top:2px">✔ Sin pendientes sin ticket · Saldo: {final_ch_saldo} s · {final_ch_saldo*95:,} kg</div>\n'
    f'      </div>\n'
    f'    </div>'
)

# ── 6. Parchear HTML ─────────────────────────────────────────────
html = HTML.read_text(encoding='utf-8')

# 6a. _md[]
old_md_pat = r'var md=window\._md\|\|\[[\s\S]*?\];'
html = re.sub(old_md_pat, new_md_block, html)
print("✅ _md[] actualizado")

# 6b. sd[] chancho por día
DATE_IDX = {'13/07':0,'14/07':1,'15/07':2,'16/07':3,'17/07':4,'18/07':5,'12/07':6}
ch_by_date = {f: (by_date[f][-1]['chSaldo'] if by_date.get(f) else 0) for f in DATE_IDX}
ts_map = {0:'al 13/07',1:'al 14/07',2:'al 15/07',3:'al 16/07',4:'al 17/07',5:'al 18/07',6:'al 12/07'}
for fecha, idx in DATE_IDX.items():
    new_s = ch_by_date.get(fecha, 0)
    ts = ts_map[idx]
    pat = rf"(Ts:'{re.escape(ts)}',Ch:')(\d+)( sacos')"
    m = re.search(pat, html)
    if m:
        html = re.sub(pat, rf"\g<1>{new_s}\3", html)
        print(f"   sd[{idx}] {fecha}: {m.group(2)}s → {new_s}s")
    else:
        print(f"   ⚠️  sd[{idx}] no encontrado")

# 6c. Balance table tbody — anclado a #sem-table-29 (no usar rfind global:
#     con el bloque de Semana 30 ya en el HTML, el "último tbody" del documento
#     ya no es necesariamente el de la tabla de Semana 29)
new_tbody = f'<tbody>\n{balance_rows}{semana_row}\n      </tbody>'
anchor29 = html.find('id="sem-table-29"')
tb_start = html.find('<tbody>', anchor29) if anchor29 != -1 else -1
tb_end   = html.find('</tbody>', tb_start) + len('</tbody>') if tb_start != -1 else -1
if tb_start != -1 and tb_end > tb_start:
    html = html[:tb_start] + new_tbody + html[tb_end:]
    print("✅ Balance tbody (Semana 29) actualizado")
else:
    print("⚠️  tbody de Semana 29 no encontrado")

# 6d. Chancho alert
old_ch_pat = r'<div style="[^"]*background:#e8f5e9;border:1\.5px solid #43a047[^"]*"[\s\S]*?</div>\s*</div>\s*</div>'
if re.search(old_ch_pat, html):
    html = re.sub(old_ch_pat, ch_alert, html, count=1)
    print("✅ Chancho alert actualizado")
else:
    print("⚠️  Chancho alert no encontrado")

# 6e. KPI card chancho (saldo en tarjeta semanal)
old_ch_kpi = r'(<div style="flex:1"><div style="color:#4e342e">Saldo</div><div style="color:#4e342e;font-size:\.95rem">)\d+( s</div></div>)'
if re.search(old_ch_kpi, html):
    html = re.sub(old_ch_kpi, rf'\g<1>{final_ch_saldo}\2', html)
    print(f"✅ KPI saldo chancho → {final_ch_saldo} s")

# ── 6f. Parches DIARIO para Semana 30 ───────────────────────────
dom19 = by_date.get('19/07', [])
if dom19:
    d19 = dom19[-1]  # último pedido del día (cierre)
    # Saldo anterior = cierre de Sáb18
    ant19 = day_closing('18/07') or {}
    saldoAntA = ant19.get('saldoA', 0)
    saldoAntP = ant19.get('saldoP', 0)
    saldoAntC = ant19.get('saldoC', 0)
    saldoAntM = ant19.get('saldoM', 0)
    saldoAntTot = ant19.get('saldoTot', 0)

    def jfmt(j, kg_per=None):
        if j == 0:
            return f'<td style="color:#aaa">0 j</td>'
        s = f'{j} j'
        if kg_per:
            s += f' <span class="kg-sub">{j*kg_per:,} kg</span>'
        return f'<td>{s}</td>'

    def jfmt_cierre(j, kg_per=None):
        if j == 0:
            return f'<td style="color:#aaa">0 j</td>'
        s = f'{j} j'
        if kg_per:
            s += f' <span class="kg-sub">{j*kg_per:,} kg</span>'
        return f'<td>{s}</td>'

    # Calcular kg de despacho y saldo
    despAkg = d19['despA'] * 15
    despPkg = d19['despP'] * 18
    despCkg = d19['despC'] * 17
    despMkg = d19['despM'] * 17
    saldoAkg = d19['saldoA'] * 15
    saldoMkg = d19['saldoM'] * 17
    saldoAntAkg = saldoAntA * 15
    saldoAntMkg = saldoAntM * 17

    def row19(cal, sal_ant, sal_ant_kg, sel, sel_kg_per, desp, desp_kg, cierre, cierre_kg):
        ant = f'{sal_ant} j <span class="kg-sub">{sal_ant_kg:,} kg</span>' if sal_ant > 0 else '<span style="color:#aaa">0 j</span>'
        s   = f'{sel} j <span class="kg-sub">{sel*sel_kg_per:,} kg</span>' if sel > 0 else '<span style="color:#aaa">0 j</span>'
        d   = f'{desp} j <span class="kg-sub">{desp_kg:,} kg</span>' if desp > 0 else '<span style="color:#aaa">0 j</span>'
        c   = f'{cierre} j <span class="kg-sub">{cierre_kg:,} kg</span>' if cierre > 0 else '<span style="color:#aaa">0 j</span>'
        return f'<tr><td>{cal}</td><td>{ant}</td><td>{s}</td><td>{d}</td><td>{c}</td></tr>\n'

    pA19 = d19['prodA']; pP19 = d19['prodP']; pC19 = d19['prodC']; pM19 = d19['prodM']
    dA19 = d19['despA']; dP19 = d19['despP']; dC19 = d19['despC']; dM19 = d19['despM']
    sA19 = d19['saldoA']; sP19 = d19['saldoP']; sC19 = d19['saldoC']; sM19 = d19['saldoM']
    ch19 = d19['chProd']; chS19 = d19['chSaldo']
    merkg19 = d19['mermaKg']; merpct19 = d19['mermaPct']
    guia19 = d19.get('guia', 'G-2998')
    ingr_s19 = d19['ingr_s']

    mer_label = '⬆️ Sobrante' if merkg19 < 0 else '📉 Merma'
    mer_color = '#e65100' if merkg19 < 0 else '#2e7d32'

    panel19_html = f'''<div class="plt-day-panel" id="day-panel-7">
<div style="background:#e8f5e9;border-left:4px solid #2e7d32;border-radius:0 8px 8px 0;padding:10px 14px">
  <div style="font-size:.72rem;font-weight:800;color:#2e7d32;margin-bottom:8px">📦 SELECCIÓN + 🚛 DESPACHO · DOM 19/07 · {ingr_s19} SACOS INGRESADOS · {guia19} · INICIO SEM. 30</div>
  <table class="plt-bal-tbl" style="background:transparent">
    <thead><tr>
      <th>Calidad</th><th>Saldo ant.</th><th>Selección</th><th>Despacho</th><th>Saldo cierre</th>
    </tr></thead>
    <tbody>
      {row19('A', saldoAntA, saldoAntAkg, pA19, 15, dA19, despAkg, sA19, saldoAkg)}      {row19('P', saldoAntP, 0, pP19, 18, dP19, despPkg, sP19, sP19*18)}      {row19('C', saldoAntC, 0, pC19, 17, dC19, despCkg, sC19, sC19*17)}      {row19('M', saldoAntM, saldoAntMkg, pM19, 17, dM19, despMkg, sM19, saldoMkg)}      <tr class="tot-row"><td>TOTAL</td><td>{saldoAntTot} j</td><td>{d19['prodTot']} j</td><td>{d19['despTotJ']} j <span class="kg-sub">{d19['despTotKg']:,} kg</span></td><td>{d19['saldoTot']} j</td></tr>
      <tr class="ch-row"><td>Chancho</td><td style="color:#aaa">— s</td><td>{ch19} s <span class="kg-sub">{ch19*95:,.0f} kg</span></td><td style="color:#aaa">—</td><td>{chS19} s</td></tr>
    </tbody>
  </table>
  <div style="margin-top:8px;font-size:.7rem;color:#555;background:#fff;border-radius:6px;padding:6px 10px">
    <strong>Guía {guia19}</strong> · A={sA19}j ({saldoAkg:,} kg) + M={sM19}j ({saldoMkg:,} kg) = {sA19*15+sM19*17:,} kg saldo &nbsp;·&nbsp;
    <span style="color:{mer_color};font-weight:700">{mer_label}: {merkg19:+.0f} kg ({merpct19:.1f}%)</span>
  </div>
</div>
            </div><!-- day-panel-7 -->'''

    # Insertar o reemplazar day-panel-7
    if 'id="day-panel-7"' in html:
        p7_start = html.find('class="plt-day-panel"', html.find('id="day-panel-7"') - 40)
        p7_start = html.rfind('<div', 0, p7_start + 1)
        p7_end   = html.find('<!-- day-panel-7 -->') + len('<!-- day-panel-7 -->')
        html = html[:p7_start] + panel19_html + html[p7_end:]
        print("✅ day-panel-7 actualizado")
    else:
        # Insertar después del comentario de cierre de panel-6
        anchor = '<!-- day-panel-6 -->'
        pos = html.find(anchor)
        if pos != -1:
            html = html[:pos + len(anchor)] + '\n\n            <!-- ── Dom 19 Jul ── -->\n            ' + panel19_html + html[pos + len(anchor):]
            print("✅ day-panel-7 insertado")
        else:
            print("⚠️  No se encontró anchor de day-panel-6")

    # ── 6g. Opción Dom 19 en el <select> ──────────────────────
    select_opt_dom19 = '<option value="7">Dom 19 Jul 🆕</option>'
    if 'value="7"' in html:
        # Actualizar etiqueta si cambió
        html = re.sub(r'<option value="7">[^<]*</option>', select_opt_dom19, html)
        print("✅ Select option Dom 19 ya existía (actualizado)")
    else:
        # Insertar ANTES del primer <option> del select
        html = re.sub(
            r'(<select class="plt-day-select"[^>]*>)\s*(<option)',
            rf'\1\n          {select_opt_dom19}\n          \2',
            html, count=1
        )
        print("✅ Select option Dom 19 insertada")

    # ── 6h. _sd entry para Dom 19 (stock KPI) ─────────────────
    sd19_ch = chS19
    sd19_entry = (
        f"{{A:'{sA19} jabas',As:'≈ {saldoAkg:,} kg<br>15 kg/jaba',"
        f"P:'0 jabas',Ps:'0 kg<br>18 kg/jaba',"
        f"C:'0 jabas',Cs:'0 kg<br>17 kg/jaba',"
        f"M:'{sM19} jabas',Ms:'≈ {saldoMkg:,} kg<br>17 kg/jaba',"
        f"T:'{d19['saldoTot']} jabas',Ts:'al 19/07',"
        f"Ch:'{sd19_ch} sacos',Chs:'al 19/07 · Sem. 30'}}"
    )
    if "Ts:'al 19/07'" in html:
        html = re.sub(r"Ts:'al 19/07',Ch:'[\d.]+( sacos')", rf"Ts:'al 19/07',Ch:'{sd19_ch}\1", html)
        print(f"✅ _sd Dom 19 chancho → {sd19_ch} sacos")
    else:
        # Insertar nueva entrada _sd después de Dom 12 (última entrada)
        pat_last_sd = r"(Chs:'al 12/07 · Sem\. 29'\})\s*(\]\)\[idx\])"
        if re.search(pat_last_sd, html):
            html = re.sub(pat_last_sd, rf"\1,\n    {sd19_entry}\n  \2", html)
            print(f"✅ _sd Dom 19 insertado: {sd19_ch} sacos")
        else:
            print("⚠️  No se encontró anchor de _sd para insertar Dom 19")

    # ── 6i. _infos entry para Dom 19 ──────────────────────────
    info19 = (
        f"Guía {guia19} &nbsp;·&nbsp; "
        f"<span style='color:#1b5e20;font-size:.95rem;font-weight:800'>"
        f"📦 {ingr_s19} sacos ingresados · Sem. 30</span>"
    )
    if '"Guía G-2998' in html or '"Guía ' + guia19 in html:
        # Reemplazar entrada existente de Dom 19
        html = re.sub(
            r'"Guía ' + re.escape(guia19) + r'[^"]*"',
            '"' + info19 + '"',
            html, count=1
        )
        print("✅ _infos Dom 19 actualizado")
    else:
        # Insertar después de la última entrada (Dom 12)
        pat_last_info = r'(Guía EG07-2977[^"]*")\s*(\];)'
        if re.search(pat_last_info, html):
            html = re.sub(pat_last_info, rf'\1,\n    "{info19}"\n  \2', html)
            print("✅ _infos Dom 19 insertado")
        else:
            print("⚠️  No se encontró anchor de _infos para insertar Dom 19")

# ── 6j. Paneles DIARIO para cualquier fecha nueva después del 19/07 ──
# (generaliza lo que 6f-6i hicieron a mano solo para el Dom 19; así cada
# fecha nueva de Semana 30 que tenga datos entra sola al tab DIARIO)
def nfmt(v):
    try: v = float(v)
    except: return str(v)
    return f'{v:g}' if v != int(v) else f'{int(v)}'

def day_summary(fecha):
    """Agrega todas las filas de un mismo día (ej: procesado + pendiente) en un solo resumen."""
    ps = by_date.get(fecha, [])
    if not ps: return None
    last = ps[-1]
    ingr_s   = sum(p['ingr_s']  for p in ps)
    ingr_kg  = sum(p['ingr_kg'] for p in ps)
    prodA = sum(p['prodA'] for p in ps); prodP = sum(p['prodP'] for p in ps)
    prodC = sum(p['prodC'] for p in ps); prodM = sum(p['prodM'] for p in ps)
    prodTot = prodA+prodP+prodC+prodM
    despA = sum(p['despA'] for p in ps); despP = sum(p['despP'] for p in ps)
    despC = sum(p['despC'] for p in ps); despM = sum(p['despM'] for p in ps)
    despTotJ  = despA+despP+despC+despM
    despTotKg = despA*15+despP*18+despC*17+despM*17
    chProd  = sum(p['chProd']  for p in ps)
    chVenta = sum(p['chVenta'] for p in ps)
    mermaKg = sum(p['mermaKg'] for p in ps if not p['desp_only'])
    guias = [p['guia'] for p in ps if p['guia'] and p['guia'] != '—']
    # Sacos comprados que ESE mismo día quedaron sin procesar (fila "pendiente":
    # tiene ingreso pero no selección ni despacho todavía) — no se deben mezclar
    # con lo que sí se procesó, para no dar la impresión de que todo se clasificó.
    pendiente_s  = sum(p['ingr_s']  for p in ps if p['desp_only'] and p['ingr_s'] > 0)
    pendiente_kg = sum(p['ingr_kg'] for p in ps if p['desp_only'] and p['ingr_s'] > 0)
    procesado_s  = ingr_s - pendiente_s
    procesado_kg = ingr_kg - pendiente_kg
    return {
        'fecha': fecha, 'dia': last['dia'],
        'ingr_s': ingr_s, 'ingr_kg': ingr_kg,
        'procesado_s': procesado_s, 'procesado_kg': procesado_kg,
        'pendiente_s': pendiente_s, 'pendiente_kg': pendiente_kg,
        'prodA': prodA, 'prodP': prodP, 'prodC': prodC, 'prodM': prodM, 'prodTot': prodTot,
        'despA': despA, 'despP': despP, 'despC': despC, 'despM': despM,
        'despTotJ': despTotJ, 'despTotKg': int(despTotKg),
        'saldoA': last['saldoA'], 'saldoP': last['saldoP'], 'saldoC': last['saldoC'], 'saldoM': last['saldoM'],
        'saldoTot': last['saldoTot'],
        'chProd': chProd, 'chVenta': chVenta, 'chSaldo': last['chSaldo'],
        'guia': guias[0] if guias else '—',
        'mermaKg': mermaKg,
        # % de merma sobre lo que SÍ se procesó ese día, no sobre el total comprado
        # (si no, los sacos pendientes de procesar inflarían el denominador y la
        # merma real se vería más chica de lo que es).
        'mermaPct': (mermaKg/procesado_kg*100) if procesado_kg else 0,
    }

DIA_LBL = {'12/07':'Dom','13/07':'Lun','14/07':'Mar','15/07':'Mié','16/07':'Jue','17/07':'Vie','18/07':'Sáb',
           '19/07':'Dom','20/07':'Lun','21/07':'Mar','22/07':'Mié','23/07':'Jue','24/07':'Vie','25/07':'Sáb'}

extra_dates = [d for d in SEM30_DATES[1:] if by_date.get(d)]  # fechas después del 19/07 con datos
if extra_dates:
    for i, fecha in enumerate(extra_dates):
        idx = 8 + i
        ds = day_summary(fecha)
        prev_date = ALL_DATES[ALL_DATES.index(fecha) - 1]
        prev = day_closing(prev_date) or {}
        saA, saP, saC, saM = prev.get('saldoA',0), prev.get('saldoP',0), prev.get('saldoC',0), prev.get('saldoM',0)
        saTot = prev.get('saldoTot', 0)
        saAkg, saMkg = saA*15, saM*17
        despAkg = ds['despA']*15; despPkg = ds['despP']*18; despCkg = ds['despC']*17; despMkg = ds['despM']*17
        saldAkg = ds['saldoA']*15; saldMkg = ds['saldoM']*17
        diaLbl = DIA_LBL.get(fecha, ds['dia'] or '')

        def r(cal, ant, ant_kg, sel, sel_kg_per, desp, desp_kg, cierre, cierre_kg):
            a = f'{ant} j <span class="kg-sub">{ant_kg:,} kg</span>' if ant > 0 else '<span style="color:#aaa">0 j</span>'
            s = f'{sel} j <span class="kg-sub">{sel*sel_kg_per:,} kg</span>' if sel > 0 else '<span style="color:#aaa">0 j</span>'
            d = f'{desp} j <span class="kg-sub">{desp_kg:,} kg</span>' if desp > 0 else '<span style="color:#aaa">0 j</span>'
            c = f'{cierre} j <span class="kg-sub">{cierre_kg:,} kg</span>' if cierre > 0 else '<span style="color:#aaa">0 j</span>'
            return f'<tr><td>{cal}</td><td>{a}</td><td>{s}</td><td>{d}</td><td>{c}</td></tr>\n'

        mer_label = '⬆️ Sobrante' if ds['mermaKg'] < 0 else '📉 Merma'
        mer_color = '#e65100' if ds['mermaKg'] < 0 else '#2e7d32'
        titulo = f"{diaLbl.upper()} {fecha}" if ds['ingr_s'] else f"{diaLbl.upper()} {fecha} · sin ingreso"

        # Encabezado: si parte de lo comprado hoy quedó sin procesar, decirlo explícito
        # en vez de mostrar el total como si todo se hubiera clasificado.
        if ds['pendiente_s'] > 0:
            ingreso_txt = f"{ds['ingr_s']} SACOS COMPRADOS ({nfmt(ds['procesado_s'])} PROCESADOS + {nfmt(ds['pendiente_s'])} PENDIENTES)"
            pendiente_note = (
                f'\n  <div style="margin-top:8px;display:flex;align-items:flex-start;gap:8px;background:#fff3e0;border:1.5px solid #ffb74d;border-radius:8px;padding:8px 12px">'
                f'\n    <span>⏳</span>'
                f'\n    <div style="font-size:.72rem;color:#795548"><strong>{nfmt(ds["pendiente_s"])} sacos ({ds["pendiente_kg"]:,.0f} kg) aún sin procesar.</strong> '
                f'No se cuentan en la selección/despacho de hoy ni en la merma — se sumarán al día que se procesen.</div>'
                f'\n  </div>'
            )
        else:
            ingreso_txt = f"{ds['ingr_s']} SACOS INGRESADOS"
            pendiente_note = ''

        panel_html = f'''<div class="plt-day-panel" id="day-panel-{idx}">
<div style="background:#e8f5e9;border-left:4px solid #2e7d32;border-radius:0 8px 8px 0;padding:10px 14px">
  <div style="font-size:.72rem;font-weight:800;color:#2e7d32;margin-bottom:8px">📦 SELECCIÓN + 🚛 DESPACHO · {titulo} · {ingreso_txt} · {ds['guia']} · SEM. 30</div>
  <table class="plt-bal-tbl" style="background:transparent">
    <thead><tr>
      <th>Calidad</th><th>Saldo ant.</th><th>Selección</th><th>Despacho</th><th>Saldo cierre</th>
    </tr></thead>
    <tbody>
      {r('A', saA, saAkg, ds['prodA'], 15, ds['despA'], despAkg, ds['saldoA'], saldAkg)}      {r('P', saP, 0, ds['prodP'], 18, ds['despP'], despPkg, ds['saldoP'], ds['saldoP']*18)}      {r('C', saC, 0, ds['prodC'], 17, ds['despC'], despCkg, ds['saldoC'], ds['saldoC']*17)}      {r('M', saM, saMkg, ds['prodM'], 17, ds['despM'], despMkg, ds['saldoM'], saldMkg)}      <tr class="tot-row"><td>TOTAL</td><td>{saTot} j</td><td>{ds['prodTot']} j</td><td>{ds['despTotJ']} j <span class="kg-sub">{ds['despTotKg']:,} kg</span></td><td>{ds['saldoTot']} j</td></tr>
      <tr class="ch-row"><td>Chancho</td><td style="color:#aaa">— s</td><td>{nfmt(ds['chProd'])} s <span class="kg-sub">{ds['chProd']*95:,.0f} kg</span></td><td style="color:#aaa">—</td><td>{ds['chSaldo']} s</td></tr>
    </tbody>
  </table>
  <div style="margin-top:8px;font-size:.7rem;color:#555;background:#fff;border-radius:6px;padding:6px 10px">
    <strong>Guía {ds['guia']}</strong> · A={ds['saldoA']}j ({saldAkg:,} kg) + M={ds['saldoM']}j ({saldMkg:,} kg) = {ds['saldoA']*15+ds['saldoM']*17:,} kg saldo &nbsp;·&nbsp;
    <span style="color:{mer_color};font-weight:700">{mer_label}: {ds['mermaKg']:+.0f} kg ({ds['mermaPct']:.1f}%)</span>
  </div>{pendiente_note}
</div>
            </div><!-- day-panel-{idx} -->'''

        # Insertar o reemplazar el panel
        marker = f'id="day-panel-{idx}"'
        if marker in html:
            p_start = html.find('class="plt-day-panel"', html.find(marker) - 40)
            p_start = html.rfind('<div', 0, p_start + 1)
            p_end   = html.find(f'<!-- day-panel-{idx} -->') + len(f'<!-- day-panel-{idx} -->')
            html = html[:p_start] + panel_html + html[p_end:]
            print(f"✅ day-panel-{idx} ({fecha}) actualizado")
        else:
            # Anclar después del ÚLTIMO day-panel-N ya existente en el documento
            last_marker = max(re.finditer(r'<!-- day-panel-(\d+) -->', html), key=lambda m: int(m.group(1)))
            pos = last_marker.end()
            html = html[:pos] + f'\n\n            <!-- ── {DIA_LBL.get(fecha,"")} {fecha} ── -->\n            ' + panel_html + html[pos:]
            print(f"✅ day-panel-{idx} ({fecha}) insertado")

        # Select: quitar insignias 🆕 viejas y crear/actualizar la opción de esta fecha
        opt_label = f"{diaLbl} {fecha[:2]} {['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][int(fecha[3:5])-1]}"
        html = html.replace(' 🆕</option>', '</option>')  # quitar cualquier insignia vieja primero
        if f'value="{idx}"' in html:
            html = re.sub(rf'<option value="{idx}">[^<]*</option>', f'<option value="{idx}">{opt_label} 🆕</option>', html, count=1)
            print(f"✅ Select option {fecha} actualizada")
        else:
            new_opt = f'<option value="{idx}">{opt_label} 🆕</option>'
            html = re.sub(
                r'(<select class="plt-day-select"[^>]*>)\s*(<option)',
                rf'\1\n          {new_opt}\n          \2',
                html, count=1
            )
            print(f"✅ Select option {fecha} insertada")

        # _sd: KPI de stock del día — se agrega al final del arreglo
        sd_entry = (
            f"{{A:'{nfmt(ds['saldoA'])} jabas',As:'{'≈ '+format(saldAkg,',') if ds['saldoA']>0 else '0'} kg<br>15 kg/jaba',"
            f"P:'{nfmt(ds['saldoP'])} jabas',Ps:'{'≈ '+format(ds['saldoP']*18,',') if ds['saldoP']>0 else '0'} kg<br>18 kg/jaba',"
            f"C:'{nfmt(ds['saldoC'])} jabas',Cs:'{'≈ '+format(ds['saldoC']*17,',') if ds['saldoC']>0 else '0'} kg<br>17 kg/jaba',"
            f"M:'{nfmt(ds['saldoM'])} jabas',Ms:'{'≈ '+format(saldMkg,',') if ds['saldoM']>0 else '0'} kg<br>17 kg/jaba',"
            f"T:'{nfmt(ds['saldoTot'])} jabas',Ts:'al {fecha}',"
            f"Ch:'{nfmt(ds['chSaldo'])} sacos',Chs:'al {fecha} · Sem. 30'}}"
        )
        sd_anchor = "])[idx];"
        if sd_anchor in html and f"Ts:'al {fecha}'" not in html:
            html = html.replace(sd_anchor, f",\n    {sd_entry}\n  {sd_anchor}", 1)
            print(f"✅ _sd {fecha} insertado")
        elif f"Ts:'al {fecha}'" in html:
            html = re.sub(rf"Ts:'al {re.escape(fecha)}',Ch:'[\d.]+( sacos')", rf"Ts:'al {fecha}',Ch:'{nfmt(ds['chSaldo'])}\1", html)
            print(f"✅ _sd {fecha} actualizado")
        else:
            print(f"⚠️  No se encontró anchor de _sd para {fecha}")

        # _infos: barra de info del día — se agrega al final del arreglo
        if ds['pendiente_s'] > 0:
            info_extra = f" &nbsp;·&nbsp; <span style='color:#e65100;font-size:.85rem;font-weight:700'>⏳ {nfmt(ds['pendiente_s'])} pendientes de procesar</span>"
        else:
            info_extra = ''
        info_entry = (
            f"Guía {ds['guia']} &nbsp;·&nbsp; "
            f"<span style='color:#1b5e20;font-size:.95rem;font-weight:800'>"
            f"📦 {ds['ingr_s']} sacos ingresados · Sem. 30</span>{info_extra}"
        )
        info_anchor = "];\n  if(bar) bar.innerHTML=infos[idx]||'';"
        if f'"Guía {ds["guia"]}' in html and ds['guia'] != '—':
            html = re.sub(r'"Guía ' + re.escape(ds['guia']) + r'[^"]*"', '"' + info_entry + '"', html, count=1)
            print(f"✅ _infos {fecha} actualizado")
        elif info_anchor in html:
            html = html.replace(info_anchor, f',\n    "{info_entry}"\n  {info_anchor}', 1)
            print(f"✅ _infos {fecha} insertado")
        else:
            print(f"⚠️  No se encontró anchor de _infos para {fecha}")

# ── 7. SEMANA 30 completa — bloque para el Resumen Semanal ──────
def replace_div_content(html, div_id, new_inner_html):
    """Reemplaza el contenido interno de <div id="div_id">...</div>, respetando divs anidados."""
    marker = f'id="{div_id}"'
    pos = html.find(marker)
    if pos == -1:
        print(f"⚠️  No se encontró <div id=\"{div_id}\">")
        return html
    tag_start = html.rfind('<div', 0, pos)
    tag_end = html.find('>', pos) + 1
    depth = 1
    i = tag_end
    while depth > 0:
        nxt_open = html.find('<div', i)
        nxt_close = html.find('</div>', i)
        if nxt_close == -1:
            print(f"⚠️  No se encontró cierre para <div id=\"{div_id}\">")
            return html
        if nxt_open != -1 and nxt_open < nxt_close:
            depth += 1
            i = nxt_open + 4
        else:
            depth -= 1
            i = nxt_close + 6
    close_tag_start = i - 6
    return html[:tag_end] + new_inner_html + html[close_tag_start:]

def make_balance_rows_generic(peds, first_date, last_date):
    out = ''
    for i, p in enumerate(peds):
        f, dia = p['fecha'], p['dia']
        is_second = (i > 0 and peds[i-1]['fecha'] == f)

        if f == first_date:
            bg='#f3e5f5'; bdr='1px solid #e1bee7'
            dcol='#6a1b9a'; diacol='#7b1fa2'; diafw='font-weight:700;'
        elif f == last_date:
            bg='#e3f2fd'; bdr='1px solid #bbdefb'
            dcol='#1565c0'; diacol='#1565c0'; diafw='font-weight:700;'
        elif is_second:
            bg='#f0f8ff'; bdr='1px solid #eee'
            dcol='#1565c0'; diacol='#888'; diafw=''
        elif i%2==1: bg='#fafafa'; bdr='1px solid #eee'; dcol='#333'; diacol='#888'; diafw=''
        else:        bg='#fff';    bdr='1px solid #eee'; dcol='#333'; diacol='#888'; diafw=''

        row = f'        <tr style="background:{bg}">\n'
        if is_second:
            row += f'          <td style="padding:7px 10px;text-align:center;font-size:.72rem;color:#1565c0;border-bottom:{bdr};font-style:italic">↳ Ped.2</td>\n'
            row += f'          <td style="padding:7px 6px;text-align:center;color:{diacol};font-size:.75rem;{diafw}border-bottom:{bdr}">{dia}</td>'
        else:
            row += f'          <td style="padding:7px 10px;text-align:center;font-weight:800;font-size:.82rem;color:{dcol};border-bottom:{bdr}">{f}</td>\n'
            row += f'          <td style="padding:7px 6px;text-align:center;color:{diacol};font-size:.75rem;{diafw}border-bottom:{bdr}">{dia}</td>'

        row += ingr_cell(p['ingr_s'], p['ingr_kg'], bdr)
        row += sel_cells(p['prodA'],p['prodP'],p['prodC'],p['prodM'],p['prodTot'],bdr,dash=p['desp_only'])
        row += desp_cells(p['guia'],p['despA'],p['despP'],p['despC'],p['despM'],p['despTotJ'],p['despTotKg'],bdr)
        row += saldo_cells(p['saldoA'],p['saldoP'],p['saldoC'],p['saldoM'],p['saldoTot'],bdr)
        row += ch_cell(p['chProd'],p['chVenta'],p['chTicket'],p['chSaldo'],bdr,ingr_kg=p['ingr_kg'])
        if p['desp_only'] or p['ingr_s']==0:
            row += merma_cells(0,0,bdr)
        else:
            row += merma_cells(p['mermaKg'],p['mermaPct'],bdr)
        row += '\n        </tr>\n'

        if p['tras'] and p['tras'] > 0:
            row += (f'        <tr style="background:#fff8e1">'
                    f'<td colspan="3" style="padding:4px 10px;font-size:.68rem;color:#e65100;border-bottom:{bdr};font-weight:700">↳ TRASIEGO</td>'
                    f'<td colspan="5" style="padding:4px 8px;font-size:.68rem;color:#795548;border-bottom:{bdr};font-style:italic">'
                    f'−{p["tras"]} jaba(s) trasigeada(s) durante llenado</td>'
                    f'<td colspan="7" style="padding:4px 8px;font-size:.68rem;color:#555;border-bottom:{bdr}">Stock ajustado calidad A: −{p["tras"]} j (−{p["tras"]*15} kg)</td>'
                    f'<td colspan="5" style="border-bottom:{bdr}"></td>'
                    f'<td colspan="3" style="border-bottom:{bdr}"></td>'
                    f'<td colspan="2" style="border-bottom:{bdr}"></td>'
                    f'</tr>\n')
        out += row
    return out

def make_semana_row_generic(week_num, ingr_s, ingr_kg, prodA,prodP,prodC,prodM,prodTot,
                             despA,despP,despC,despM,despTotJ,despTotKg,
                             saldoA,saldoP,saldoC,saldoM,saldoTot,
                             chProd,chVenta,chSaldo,mermaKg,merma_pct):
    ch_sc = '#2e7d32' if chSaldo==0 else '#4e342e'
    ch_sb = '#c8e6c9' if chSaldo==0 else '#efebe9'
    mc = '#1b5e20' if mermaKg>=0 else '#e65100'
    ms = '+' if mermaKg>=0 else '-'
    def sc(v): return '#311b92' if v>0 else '#aaa'
    return f'''        <tr style="background:#eceff1;border-top:2px solid #90a4ae">
          <td colspan="3" style="padding:6px 10px;text-align:center;font-weight:900;font-size:.85rem;color:#263238;line-height:1.5">SEMANA {week_num}<br><span style="font-size:.67rem;font-weight:600;color:#1565c0">{ingr_s} s · {ingr_kg:,} kg ingresados</span></td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#0d47a1;background:#dbeafe">{prodA}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#0d47a1;background:#dbeafe">{prodP}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#0d47a1;background:#dbeafe">{prodC}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#0d47a1;background:#dbeafe">{prodM}</td>
          <td style="padding:8px;text-align:center;font-weight:900;color:#0d47a1;background:#bfdbfe">{prodTot:,} j</td>
          <td style="padding:8px;text-align:center;color:#bbb;font-size:.7rem;background:#fff9f0">—</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#1b5e20;background:#dcfce7">{despA}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#1b5e20;background:#dcfce7">{despP}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#1b5e20;background:#dcfce7">{despC}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:#1b5e20;background:#dcfce7">{despM}</td>
          <td style="padding:8px;text-align:center;font-weight:900;color:#1b5e20;background:#bbf7d0">{despTotJ:,} j</td>
          <td style="padding:8px;text-align:center;font-weight:900;color:#1b5e20;background:#bbf7d0">{despTotKg:,} kg</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:{sc(saldoA)};background:#ede7f6">{saldoA}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:{sc(saldoP)};background:#ede7f6">{saldoP}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:{sc(saldoC)};background:#ede7f6">{saldoC}</td>
          <td style="padding:8px;text-align:center;font-weight:800;color:{sc(saldoM)};background:#ede7f6">{saldoM}</td>
          <td style="padding:8px;text-align:center;font-weight:900;color:#311b92;background:#ddd6fe">{saldoTot} j</td>
          <td style="padding:6px 4px;text-align:center;font-weight:800;color:#4e342e;background:#efebe9;line-height:1.4"><div>{nfmt(chProd)} s prod.</div><div style="font-size:.66rem;color:#a1887f">{nfmt(chProd*95)} kg</div><div style="font-size:.6rem;font-weight:700;color:#c62828">{(chProd*95/ingr_kg*100) if ingr_kg>0 else 0:.1f}% del ingreso</div></td><td style="padding:6px 4px;text-align:center;font-weight:900;color:#e65100;background:#fff8f0;line-height:1.4"><div>{nfmt(chVenta)} s</div><div style="font-size:.66rem;color:#e57373">{nfmt(chVenta*95)} kg</div></td><td style="padding:6px 4px;text-align:center;font-weight:900;color:{ch_sc};background:{ch_sb};line-height:1.4"><div>{nfmt(chSaldo)} s</div><div style="font-size:.66rem;color:#a5d6a7">{nfmt(chSaldo*95)} kg</div></td>
          <td style="padding:8px;text-align:center;font-weight:900;color:{mc};background:#fff3e0">{mermaKg:+.0f} kg</td>
          <td style="padding:8px;text-align:center;font-weight:900;color:{mc};background:#fff3e0">{ms}{abs(merma_pct):.1f}%</td>
        </tr>'''

THEAD_TMPL = '''      <thead>
        <tr>
          <th colspan="3" style="padding:8px 10px;background:#37474f;color:#fff;text-align:center;border-radius:8px 0 0 0"><div style="font-size:.82rem;font-weight:900;letter-spacing:.04em">SEMANA {wk}</div><div style="font-size:.68rem;font-weight:400;opacity:.85;margin-top:2px">{rango}</div></th>
          <th colspan="5" style="padding:8px 10px;background:#1565c0;color:#fff;text-align:center;font-size:.72rem;letter-spacing:.05em">🌿 SELECCIÓN (jabas)</th>
          <th colspan="7" style="padding:8px 10px;background:#1b5e20;color:#fff;text-align:center;font-size:.72rem;letter-spacing:.05em">🚛 DESPACHO A SUPESA</th>
          <th colspan="5" style="padding:8px 10px;background:#4527a0;color:#fff;text-align:center;font-size:.72rem;letter-spacing:.05em">🔄 SALDO EN PLANTA</th>
          <th colspan="3" style="padding:8px 10px;background:#795548;color:#fff;text-align:center;font-size:.72rem;letter-spacing:.05em">🐷 CHANCHO</th><th colspan="3" style="padding:8px 10px;background:#b71c1c;color:#fff;text-align:center;font-size:.72rem;letter-spacing:.05em;border-radius:0 8px 0 0">⚖️ MERMA</th>
        </tr>
        <tr style="background:#f0f0f0">
          <th style="padding:6px 8px;text-align:center;color:#555;font-size:.7rem;border-bottom:2px solid #ddd">Fecha</th>
          <th style="padding:6px 8px;text-align:center;color:#555;font-size:.7rem;border-bottom:2px solid #ddd">Día</th><th style="padding:6px 8px;text-align:center;color:#1565c0;font-size:.7rem;border-bottom:2px solid #1565c0;min-width:58px;background:#e3f2fd">📦 INGRESO<br><span style="font-size:.62rem;font-weight:400">(s / kg)</span></th>
          <th style="padding:6px 8px;text-align:center;color:#1565c0;font-size:.7rem;border-bottom:2px solid #1565c0">A</th>
          <th style="padding:6px 8px;text-align:center;color:#1565c0;font-size:.7rem;border-bottom:2px solid #1565c0">P</th>
          <th style="padding:6px 8px;text-align:center;color:#1565c0;font-size:.7rem;border-bottom:2px solid #1565c0">C</th>
          <th style="padding:6px 8px;text-align:center;color:#1565c0;font-size:.7rem;border-bottom:2px solid #1565c0">M</th>
          <th style="padding:6px 8px;text-align:center;color:#1565c0;font-size:.7rem;border-bottom:2px solid #1565c0;font-weight:900">Tot j</th>
          <th style="padding:6px 8px;text-align:center;color:#1b5e20;font-size:.7rem;border-bottom:2px solid #1b5e20;min-width:54px">N° Guía</th>
          <th style="padding:6px 8px;text-align:center;color:#1b5e20;font-size:.7rem;border-bottom:2px solid #1b5e20">A</th>
          <th style="padding:6px 8px;text-align:center;color:#1b5e20;font-size:.7rem;border-bottom:2px solid #1b5e20">P</th>
          <th style="padding:6px 8px;text-align:center;color:#1b5e20;font-size:.7rem;border-bottom:2px solid #1b5e20">C</th>
          <th style="padding:6px 8px;text-align:center;color:#1b5e20;font-size:.7rem;border-bottom:2px solid #1b5e20">M</th>
          <th style="padding:6px 8px;text-align:center;color:#1b5e20;font-size:.7rem;border-bottom:2px solid #1b5e20;font-weight:900">Tot j</th>
          <th style="padding:6px 8px;text-align:center;color:#1b5e20;font-size:.7rem;border-bottom:2px solid #1b5e20;font-weight:900">Tot kg</th>
          <th style="padding:6px 8px;text-align:center;color:#4527a0;font-size:.7rem;border-bottom:2px solid #4527a0">A</th>
          <th style="padding:6px 8px;text-align:center;color:#4527a0;font-size:.7rem;border-bottom:2px solid #4527a0">P</th>
          <th style="padding:6px 8px;text-align:center;color:#4527a0;font-size:.7rem;border-bottom:2px solid #4527a0">C</th>
          <th style="padding:6px 8px;text-align:center;color:#4527a0;font-size:.7rem;border-bottom:2px solid #4527a0">M</th>
          <th style="padding:6px 8px;text-align:center;color:#4527a0;font-size:.7rem;border-bottom:2px solid #4527a0;font-weight:900">Tot j</th>
          <th style="padding:6px 8px;text-align:center;color:#795548;font-size:.7rem;border-bottom:2px solid #795548">Prod (s)</th><th style="padding:6px 8px;text-align:center;color:#e65100;font-size:.7rem;border-bottom:2px solid #e65100">💰 Venta (s)</th><th style="padding:6px 8px;text-align:center;color:#795548;font-size:.7rem;border-bottom:2px solid #795548">Saldo (s)</th>
          <th style="padding:6px 8px;text-align:center;color:#b71c1c;font-size:.7rem;border-bottom:2px solid #b71c1c">Merma kg</th>
          <th style="padding:6px 8px;text-align:center;color:#b71c1c;font-size:.7rem;border-bottom:2px solid #b71c1c">Merma %</th>
        </tr>
      </thead>'''

sem30 = [p for p in pedidos if p['fecha'] in SEM30_DATES]
dates_with_data_30 = [d for d in SEM30_DATES if by_date.get(d)]

if sem30:
    last_date_30 = dates_with_data_30[-1]
    close30 = day_closing(last_date_30) or {}
    final_saldoA_30 = close30.get('saldoA', 0)
    final_saldoP_30 = close30.get('saldoP', 0)
    final_saldoC_30 = close30.get('saldoC', 0)
    final_saldoM_30 = close30.get('saldoM', 0)
    final_saldo_tot_30 = close30.get('saldoTot', 0)
    final_ch_saldo_30 = close30.get('chSaldo', 0)

    ch_opening_30 = final_ch_saldo   # arrastre del cierre de Semana 29
    total_ingr_s_30  = sum(p['ingr_s']  for p in sem30)
    total_ingr_kg_30 = sum(p['ingr_kg'] for p in sem30)
    total_prodA_30 = sum(p['prodA'] for p in sem30)
    total_prodP_30 = sum(p['prodP'] for p in sem30)
    total_prodC_30 = sum(p['prodC'] for p in sem30)
    total_prodM_30 = sum(p['prodM'] for p in sem30)
    total_prodTot_30 = sum(p['prodTot'] for p in sem30)
    total_despA_30 = sum(p['despA'] for p in sem30)
    total_despP_30 = sum(p['despP'] for p in sem30)
    total_despC_30 = sum(p['despC'] for p in sem30)
    total_despM_30 = sum(p['despM'] for p in sem30)
    total_despTotJ_30  = sum(p['despTotJ']  for p in sem30)
    total_despTotKg_30 = sum(p['despTotKg'] for p in sem30)
    total_chProd_30  = sum(p['chProd']  for p in sem30)
    total_chVenta_30 = sum(p['chVenta'] for p in sem30)
    total_mermaKg_30 = sum(p['mermaKg'] for p in sem30 if not p['desp_only'])
    total_merma_pct_30 = (total_mermaKg_30 / total_ingr_kg_30 * 100) if total_ingr_kg_30 > 0 else 0
    new_ch_total_30 = ch_opening_30 + total_chProd_30

    ch_tickets_str_30 = ' + '.join(
        f"Tkt {p['chTicket']} ({nfmt(p['chVenta'])}s)"
        for p in sem30 if p['chVenta'] > 0 and p['chTicket']
    )

    sel_kg_30  = prod_kg(total_prodA_30, total_prodP_30, total_prodC_30, total_prodM_30)
    ch_kg_30   = total_chProd_30 * 95
    clas_kg_30 = sel_kg_30 + ch_kg_30
    pct_sel_30  = round(sel_kg_30  / total_ingr_kg_30 * 100) if total_ingr_kg_30 > 0 else 0
    pct_ch_30   = round(ch_kg_30   / total_ingr_kg_30 * 100) if total_ingr_kg_30 > 0 else 0
    pct_clas_30 = round(clas_kg_30 / total_ingr_kg_30 * 100) if total_ingr_kg_30 > 0 else 0
    merma_sign_30  = '⬆️ Sobrante' if total_mermaKg_30 < 0 else '📉 Merma'
    merma_color_30 = '#b71c1c' if total_mermaKg_30 < 0 else '#1b5e20'

    print(f"\n✅ Semana 30 (en curso, {len(sem30)} día(s) con datos: {dates_with_data_30}):")
    print(f"   Ingreso: {total_ingr_s_30}s / {total_ingr_kg_30:,}kg · Selección: {total_prodTot_30}j · Despacho: {total_despTotJ_30}j")
    print(f"   Saldo cierre ({last_date_30}): A={final_saldoA_30} P={final_saldoP_30} C={final_saldoC_30} M={final_saldoM_30} tot={final_saldo_tot_30}")
    print(f"   Chancho: aper={ch_opening_30}+prod={nfmt(total_chProd_30)}={nfmt(new_ch_total_30)} · venta={nfmt(total_chVenta_30)} · saldo={nfmt(final_ch_saldo_30)}")

    def kpi_cierre_card(label, jabas, kg, highlight):
        if highlight:
            lbl_c='#e65100'; val_c='#e65100'; sub_c='#e65100'; border='border:2px solid #e65100;'; bg=''
        elif jabas>0:
            lbl_c='#4527a0'; val_c='#4527a0'; sub_c='#888'; border=''; bg=''
        else:
            lbl_c='#aaa'; val_c='#bbb'; sub_c='#bbb'; border=''; bg='background:#f5f5f5;'
        sub = f'≈ {nfmt(kg)} kg' if jabas>0 else '— kg'
        return (f'<div class="plt-kpi" style="flex:1;{border}{bg}">'
                f'<div class="plt-kpi-lbl" style="color:{lbl_c}">{label} · CIERRE SEM 30</div>'
                f'<div class="plt-kpi-val" style="font-size:1.1rem!important;color:{val_c}">{nfmt(jabas)} jabas</div>'
                f'<div class="plt-kpi-sub" style="color:{sub_c}">{sub}</div>'
                f'</div>')

    kg_A_30 = final_saldoA_30*15; kg_P_30 = final_saldoP_30*18; kg_C_30 = final_saldoC_30*17; kg_M_30 = final_saldoM_30*17

    chancho_alert_30 = (
        f'<div style="display:flex;align-items:flex-start;gap:8px;background:#e8f5e9;border:1.5px solid #43a047;border-radius:8px;padding:10px 14px;margin-top:10px">\n'
        f'      <span>✅</span>\n'
        f'      <div>\n'
        f'        <div style="font-weight:800;color:#1b5e20;font-size:.75rem">CHANCHO · SEM. 30 (en curso)</div>\n'
        f'        <div style="font-size:.68rem;color:#388e3c;margin-top:3px">Apertura: {nfmt(ch_opening_30)} s (saldo Sem.29) · Producidos sem: {nfmt(total_chProd_30)} s · Total: {nfmt(new_ch_total_30)} s'
        + (f' · {ch_tickets_str_30}' if ch_tickets_str_30 else ' · Sin ventas registradas aún')
        + '</div>\n'
        f'        <div style="font-size:.67rem;color:#555;margin-top:2px">Saldo: {nfmt(final_ch_saldo_30)} s · {nfmt(final_ch_saldo_30*95)} kg</div>\n'
        f'      </div>\n'
        f'    </div>'
    )

    kpi_html_30 = f'''<div style="display:flex;gap:14px;align-items:flex-start;margin-bottom:16px">

    <div style="flex:1;min-width:210px;background:#f9f9f9;border:1px solid #e0e0e0;border-radius:10px;padding:12px 14px;font-size:.78rem">
      <div style="font-weight:800;color:#444;margin-bottom:10px;font-size:.72rem;text-transform:uppercase;letter-spacing:.05em">📊 Resumen Semana 30</div>
      <table style="width:100%;border-collapse:collapse">
        <tr><td style="color:#555;padding:3px 0">📦 Ingreso total ({total_ingr_s_30} s)</td><td style="text-align:right;font-weight:800;font-size:.88rem">{total_ingr_kg_30:,} kg <span style="color:#aaa;font-weight:500;font-size:.78em">100%</span></td></tr>
        <tr><td style="color:#555;padding:3px 0">🎋 Jabas seleccionadas ({total_prodTot_30:,} j)</td><td style="text-align:right;font-weight:800;font-size:.88rem">{sel_kg_30:,} kg <span style="color:#1565c0;font-weight:700;font-size:.82em">{pct_sel_30}%</span></td></tr>
        <tr><td style="color:#555;padding:3px 0">🐷 Chancho producido ({nfmt(total_chProd_30)} s)</td><td style="text-align:right;font-weight:800;font-size:.88rem">{nfmt(ch_kg_30)} kg <span style="color:#795548;font-weight:700;font-size:.82em">{pct_ch_30}%</span></td></tr>
        <tr style="border-top:1px solid #ddd"><td style="font-weight:700;color:#333;padding-top:5px">Total clasificado</td><td style="text-align:right;font-weight:900;font-size:.92rem;padding-top:5px">{nfmt(clas_kg_30)} kg <span style="color:#aaa;font-weight:500;font-size:.78em">{pct_clas_30}%</span></td></tr>
        <tr><td style="font-weight:800;color:{merma_color_30};padding-top:3px">{merma_sign_30}</td><td style="text-align:right;font-weight:900;color:{merma_color_30};font-size:.95rem">{total_mermaKg_30:+.0f} kg ({total_merma_pct_30:+.1f}%)</td></tr>
        <tr style="border-top:1px solid #e0e0e0"><td style="color:#888;padding-top:5px;font-size:.7rem">🚛 Despacho semana</td><td style="text-align:right;color:#888;font-size:.7rem;padding-top:5px">{total_despTotKg_30:,} kg</td></tr>
      </table>
      {chancho_alert_30}
      <div style="margin-top:8px;font-size:.68rem;color:#aaa">{total_ingr_s_30} sacos × 95 kg = {total_ingr_kg_30:,} · {nfmt(total_chProd_30)} s chancho × 95 = {nfmt(ch_kg_30)} · Jabas = suma diaria</div>
    </div>

    <div style="flex:2;min-width:0">
      <div style="font-size:.62rem;font-weight:800;text-transform:uppercase;letter-spacing:.08em;color:#666;margin-bottom:6px">📊 Reporte Semana 30 · 19–25 Jul 2026 (en curso)</div>
      <div style="display:flex;gap:8px;margin-bottom:8px">
        <div class="plt-kpi" style="background:#e3f2fd;flex:1">
          <div class="plt-kpi-lbl" style="color:#0d47a1">SACOS COMPRADOS</div>
          <div class="plt-kpi-val" style="color:#0d47a1;font-size:1.3rem!important">{total_ingr_s_30} sacos</div>
          <div class="plt-kpi-sub">semana 30 · a la fecha</div>
        </div>
        <div class="plt-kpi" style="background:#e8f5e9;flex:1">
          <div class="plt-kpi-lbl" style="color:#1b5e20">SELECCIÓN</div>
          <div class="plt-kpi-val" style="color:#1b5e20;font-size:1.3rem!important">{total_prodTot_30:,} jabas</div>
          <div class="plt-kpi-sub">jabas seleccionadas</div>
        </div>
        <div class="plt-kpi" style="background:#f3e5f5;flex:1">
          <div class="plt-kpi-lbl" style="color:#4a148c">DESPACHO</div>
          <div class="plt-kpi-val" style="color:#4a148c;font-size:1.3rem!important">{total_despTotKg_30:,} kg</div>
          <div class="plt-kpi-sub">{total_despTotJ_30:,} jabas enviadas</div>
        </div>
        <div class="plt-kpi" style="flex:1">
          <div class="plt-kpi-lbl">SALDO ACTUAL</div>
          <div class="plt-kpi-val" style="font-size:1.3rem!important;color:#e65100">{final_saldo_tot_30} jabas</div>
          <div class="plt-kpi-sub">saldo cierre {last_date_30}</div>
        </div>
      </div>
      <div style="display:flex;gap:8px">
        {kpi_cierre_card('A', final_saldoA_30, kg_A_30, True)}
        {kpi_cierre_card('P', final_saldoP_30, kg_P_30, False)}
        {kpi_cierre_card('C', final_saldoC_30, kg_C_30, False)}
        {kpi_cierre_card('M', final_saldoM_30, kg_M_30, False)}
        <div class="plt-kpi" style="background:#fff8f0;flex:1.2">
          <div class="plt-kpi-lbl" style="color:#4e342e">🐷 CHANCHO</div>
          <div style="display:flex;gap:6px;margin-top:4px;font-size:.72rem;font-weight:700;text-align:center">
            <div style="flex:1"><div style="color:#795548">Prod</div><div style="color:#4e342e;font-size:.95rem">{nfmt(total_chProd_30)} s</div></div>
            <div style="flex:1"><div style="color:#e65100">Venta</div><div style="color:#e65100;font-size:.95rem">{nfmt(total_chVenta_30)} s</div></div>
            <div style="flex:1"><div style="color:#4e342e">Saldo</div><div style="color:#4e342e;font-size:.95rem">{nfmt(final_ch_saldo_30)} s</div></div>
          </div>
        </div>
      </div>
    </div>

  </div>'''

    balance_rows_30 = make_balance_rows_generic(sem30, sem30[0]['fecha'], last_date_30)
    semana_row_30 = make_semana_row_generic(30, total_ingr_s_30, total_ingr_kg_30,
        total_prodA_30,total_prodP_30,total_prodC_30,total_prodM_30,total_prodTot_30,
        total_despA_30,total_despP_30,total_despC_30,total_despM_30,total_despTotJ_30,total_despTotKg_30,
        final_saldoA_30,final_saldoP_30,final_saldoC_30,final_saldoM_30,final_saldo_tot_30,
        total_chProd_30,total_chVenta_30,final_ch_saldo_30,total_mermaKg_30,total_merma_pct_30)

    table_html_30 = f'''<div class="plt-sec" style="padding:14px 16px">
    <h3 style="margin-bottom:14px">📆 Balance Diario · Semana 30 · 19–25 Jul 2026 (en curso)</h3>

    <div style="overflow-x:auto">
    <table style="width:100%;border-collapse:collapse;font-size:.78rem;min-width:860px">
{THEAD_TMPL.format(wk=30, rango='Dom 19 – Sáb 25 Jul 2026')}
      <tbody>
{balance_rows_30}{semana_row_30}
      </tbody>
    </table>
    </div>

    <div class="plt-note" style="margin-top:10px">
      ⚖️ Merma = Ingreso (sacos × 95 kg) − Jabas seleccionadas (kg) − Chancho (95 kg/saco) &nbsp;·&nbsp; Semana en curso — la tabla se completa a medida que registras más días en Balance_Semanal.xlsx
      &nbsp;·&nbsp; Merma negativa = sobrante (stock arrastrado) &nbsp;·&nbsp;
      Conversión: A=15 · P=18 · C=17 · M=17 kg/jaba
    </div>
  </div>'''

    html = replace_div_content(html, 'sem-kpis-30', kpi_html_30)
    html = replace_div_content(html, 'sem-table-30', table_html_30)
    print("✅ Bloque 'Resumen Semanal' de Semana 30 generado e insertado")
else:
    print("ℹ️  Semana 30: todavía no hay filas con fecha en Balance_Semanal.xlsx — nada que mostrar en el Resumen")

HTML.write_text(html, encoding='utf-8')
print(f"\n✅ Listo — {len(sem29)} pedidos Sem29 + {len(pedidos)-len(sem29)} pedidos Sem30 · chancho saldo final: {final_ch_saldo} s · {final_ch_saldo*95} kg")
