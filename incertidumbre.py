"""Análisis de incertidumbre - Encuesta de prácticas de seguridad en desarrollo ágil.
Uso: python3 incertidumbre.py Extended_Survey_Results.xlsm
"""
import sys, numpy as np, openpyxl
from scipy.stats import binomtest, fisher_exact
from statsmodels.stats.proportion import proportion_confint
from statsmodels.stats.contingency_tables import mcnemar, cochrans_q

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
rows = list(wb['Complete Responses'].iter_rows(values_only=True))
H, D = rows[0], [r for r in rows[1:] if r[0] is not None]
ID_CORTE = 91  # IDs <= 91 = 24 respuestas de la 1.a iteración (hoja "Other Graphics")
ola = np.array([1 if r[0] <= ID_CORTE else 2 for r in D])

fases = {'Pre-proyecto':31,'Requisitos':32,'Diseño':33,'Implementación':34,
         'Pruebas':35,'Release':36,'Operaciones':37}
X = np.array([[1 if r[c]=='Yes' else 0 for c in fases.values()] for r in D])
N = len(D)
print(f"N total={N}  (iteración 1: {sum(ola==1)}, nuevas: {sum(ola==2)})\n")

print("1) Proporción por fase con IC 95% (Wilson)")
print(f"{'Fase':16}{'Iter.1 (n=24)':>22}{'Nuevas (n=10)':>22}{'Total (n=34)':>24}")
for j,f in enumerate(fases):
    out=[]
    for m in (ola==1, ola==2, np.ones(N,bool)):
        k,n=X[m,j].sum(),m.sum(); lo,hi=proportion_confint(k,n,method='wilson')
        out.append(f"{k}/{n} {k/n:4.0%} [{lo:3.0%}-{hi:3.0%}]")
    print(f"{f:16}{out[0]:>22}{out[1]:>22}{out[2]:>24}")

j_op = list(fases).index('Operaciones')
print("\n2) ¿Las fases difieren entre sí? Cochran's Q (respuestas pareadas)")
q = cochrans_q(X); print(f"   Q={q.statistic:.2f}, gl={len(fases)-1}, p={q.pvalue:.4f}")

print("\n3) Operaciones vs. cada fase - McNemar exacto (mismas personas), Holm")
res=[]
for j,f in enumerate(fases):
    if j==j_op: continue
    b=int(((X[:,j_op]==1)&(X[:,j]==0)).sum()); c=int(((X[:,j_op]==0)&(X[:,j]==1)).sum())
    res.append((f,b,c,mcnemar([[0,b],[c,0]],exact=True).pvalue))
ps=sorted(range(len(res)),key=lambda i:res[i][3]); adj=[0]*len(res); run=0
for r,i in enumerate(ps): run=max(run,min(1,(len(res)-r)*res[i][3])); adj[i]=run
for (f,b,c,p),pa in zip(res,adj):
    print(f"   {f:16} solo Op={b:2d}  solo {f[:5]}={c:2d}  p={p:.3f}  p_Holm={pa:.3f}")

print("\n4) Bootstrap (10 000 remuestreos): ¿qué tan estable es el ranking?")
rng=np.random.default_rng(2026); B=10000; lowest=np.zeros(len(fases)); diffs=[]
for _ in range(B):
    s=X[rng.integers(0,N,N)].mean(0)
    m=s.min(); w=np.flatnonzero(np.isclose(s,m)); lowest[w]+=1/len(w)
    diffs.append(s[j_op]-np.delete(s,j_op).min())
for f,p in zip(fases,lowest/B): print(f"   P({f} es la más baja) = {p:.2f}")
lo,hi=np.percentile(diffs,[2.5,97.5])
print(f"   Brecha Operaciones - siguiente fase más baja: IC95% [{lo:+.0%}, {hi:+.0%}]")

print("\n5) ¿Cambió Operaciones entre olas? Fisher exacto")
a=X[ola==1,j_op].sum(); b=X[ola==2,j_op].sum()
p=fisher_exact([[a,24-a],[b,10-b]]).pvalue
print(f"   Iter.1 {a}/24 vs nuevas {b}/10 -> p={p:.3f}")

print("\n6) Sensibilidad a 'Don't know / No answer' (columnas de impacto, por fase)")
print("   Se recalcula '% impacto alto o muy alto' con tres supuestos sobre las respuestas NS/NR")
for f,kw in [('Requisitos','REQUIREMENTS'),('Diseño','DESIGN'),('Implementación','IMPLEMENTATION'),
             ('V&V','VERIFICATION'),('Release','RELEASE')]:
    cols=[c for c,h in enumerate(H) if h and kw in str(h) and 'performed' in str(h)][1::2]  # impares = impacto
    for nombre,G in (('Iter.1',[r for r,o in zip(D,ola) if o==1]),('Total',D)):
        v=[str(r[c]) for r in G for c in cols if r[c] not in (None,'N/A')]
        dk=sum('know' in x for x in v); hi_=sum(x in ('High','Very high') for x in v); n=len(v)
        print(f"   {f:15}{nombre:7} NS/NR={dk/n:4.0%} | alto: excluyendo NS/NR={hi_/(n-dk):4.0%}, "
              f"NS/NR=no alto={hi_/n:4.0%}, NS/NR=alto={(hi_+dk)/n:4.0%}")
