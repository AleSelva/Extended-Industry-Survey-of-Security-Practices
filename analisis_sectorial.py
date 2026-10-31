"""Análisis de la composición sectorial - Encuesta de prácticas de seguridad en desarrollo ágil.

Compara el sector de las organizaciones entre la iteración 1 (24 respuestas) y las
respuestas nuevas (10), y prueba si la integración de seguridad por fase del SDLC
depende del sector (Software y TI frente a otros sectores).

Uso:     python3 analisis_sectorial.py Extended_Survey_Results.xlsm
Requiere: openpyxl, scipy
"""
import sys
import openpyxl
from scipy.stats import fisher_exact

ID_CORTE = 91                     # Response ID <= 91: 24 respuestas de la iteración 1 (hoja "Other Graphics")
COL_SECTOR = 15                   # columna P de "Complete Responses" (índice desde 0)
SW = 'Software development and IT services'
FASES = {'Pre-proyecto': 31, 'Requisitos': 32, 'Diseño': 33, 'Implementación': 34,
         'Pruebas': 35, 'Release': 36, 'Operaciones': 37}   # columnas AF:AL
SECTORES = [SW, 'Education', 'Banking', 'Insurance and financial services', 'Public administration',
            'Construction and transportation', 'Consulting', 'Agriculture, forestry and fishing',
            'Manufacturing', 'Tourism']


def holm(pvals):
    """Valores p ajustados por el método de Holm (en el mismo orden de entrada)."""
    m = len(pvals)
    orden = sorted(range(m), key=lambda i: pvals[i])
    ajust, acumulado = [0.0] * m, 0.0
    for rango, i in enumerate(orden):
        acumulado = max(acumulado, min(1.0, (m - rango) * pvals[i]))
        ajust[i] = acumulado
    return ajust


def fisher(si_a, n_a, si_b, n_b):
    """Prueba exacta de Fisher bilateral para dos proporciones independientes."""
    return fisher_exact([[si_a, n_a - si_a], [si_b, n_b - si_b]]).pvalue


def main(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    filas = list(wb['Complete Responses'].iter_rows(values_only=True))
    datos = [r for r in filas[1:] if r[0] is not None]
    it1 = [r for r in datos if r[0] <= ID_CORTE]
    nuevas = [r for r in datos if r[0] > ID_CORTE]
    es_sw = lambda r: r[COL_SECTOR] == SW
    print(f"N = {len(datos)}  (iteración 1: {len(it1)}, nuevas: {len(nuevas)})\n")

    # 1. Composición por sector
    print("1) Respuestas por sector")
    print(f"{'Sector':38}{'Iter. 1':>14}{'Nuevas':>14}{'Total':>14}")
    for s in SECTORES:
        a = sum(r[COL_SECTOR] == s for r in it1)
        b = sum(r[COL_SECTOR] == s for r in nuevas)
        print(f"{s:38}{a:>5} ({a/len(it1):5.1%}){b:>5} ({b/len(nuevas):5.1%}){a+b:>5} ({(a+b)/len(datos):5.1%})")

    # 2. Software y TI frente a otros sectores, entre iteraciones
    a, b = sum(map(es_sw, it1)), sum(map(es_sw, nuevas))
    p = fisher(a, len(it1), b, len(nuevas))
    print(f"\n2) Software y TI: iteración 1 = {a}/{len(it1)} ({a/len(it1):.0%}), "
          f"nuevas = {b}/{len(nuevas)} ({b/len(nuevas):.0%})  ->  Fisher p = {p:.3f}")

    # 3. Integración de seguridad por fase según grupo sectorial
    sw = [r for r in datos if es_sw(r)]
    otros = [r for r in datos if not es_sw(r)]
    print(f"\n3) Integración por fase: Software y TI (n = {len(sw)}) frente a otros sectores (n = {len(otros)})")
    print(f"{'Fase':16}{'SW y TI':>16}{'Otros':>16}{'Dif. (pts)':>12}{'p Fisher':>11}{'p Holm':>9}")
    res = []
    for fase, c in FASES.items():
        x1 = sum(r[c] == 'Yes' for r in sw)
        x2 = sum(r[c] == 'Yes' for r in otros)
        res.append((fase, x1, x2, fisher(x1, len(sw), x2, len(otros))))
    for (fase, x1, x2, p), pa in zip(res, holm([t[3] for t in res])):
        p1, p2 = x1 / len(sw), x2 / len(otros)
        print(f"{fase:16}{x1:>5}/{len(sw)} ({p1:4.0%}){x2:>5}/{len(otros)} ({p2:4.0%})"
              f"{(p1-p2)*100:>11.0f}{p:>11.4f}{pa:>9.4f}")

    # 4. Requisitos por iteración y grupo sectorial
    req = FASES['Requisitos']
    print("\n4) Requisitos: 'Sí' por iteración y grupo sectorial")
    for nombre, grupo in (('Software y TI', True), ('Otros sectores', False)):
        g1 = [r for r in it1 if es_sw(r) == grupo]
        g2 = [r for r in nuevas if es_sw(r) == grupo]
        s1, s2 = sum(r[req] == 'Yes' for r in g1), sum(r[req] == 'Yes' for r in g2)
        print(f"   {nombre:15} iteración 1: {s1}/{len(g1)} ({s1/len(g1):.0%})   "
              f"nuevas: {s2}/{len(g2)} ({s2/len(g2):.0%})")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit("Uso: python3 analisis_sectorial.py Extended_Survey_Results.xlsm")
    main(sys.argv[1])
