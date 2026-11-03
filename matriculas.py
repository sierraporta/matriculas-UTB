#!/usr/bin/env python3
"""Monitoreo de matriculas de primera vez (reportes BAN, Retencion Levantada).

Uso, desde la carpeta del proyecto:
    python3 matriculas.py                 # procesa todo lo que haya en la carpeta
    python3 matriculas.py --fecha-reporte 2026-11-03   # fecha que aparece en la portada (por defecto, hoy)
    python3 matriculas.py --por-lamina 4  # programas por diapositiva en cada escuela (por defecto 2)
    python3 matriculas.py --bruto         # reproduce el conteo antiguo (sin quitar devoluciones ni duplicados)
    python3 matriculas.py --anios 3 --salida Salida

Entradas:  BAN - Reporte de Retención Levantada AAAAPP.xls  (TSV con encabezado, PP = 10 o 20)
           metas.csv  (anio, sem, programa, escuela, meta)
Salidas :  Salida/Matriculas_AAAAMMDD.pptx (documento con diapositivas), Matriculas_resumen.xlsx, resumen.txt
           --figuras guarda ademas las figuras sueltas en Salida/Figuras
Requiere: pandas, matplotlib, openpyxl, python-pptx (pip install python-pptx openpyxl)
Solo se escriben agregados; ningun dato personal sale de la memoria del script.

Definicion de matriculado nuevo: Tipo_estudiante == PRIMERA VEZ, sin Devolucion == Si,
una fila por (Id, Programa) conservando el ultimo pago.
"""
import argparse, re, shutil, sys, tempfile
from pathlib import Path
import numpy as np, pandas as pd, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.colors import to_hex, to_rgb
from matplotlib.lines import Line2D

if any("Lato" in f.name for f in fm.fontManager.ttflist):
    plt.rcParams["font.family"] = "Lato"
plt.rcParams.update({"axes.spines.top": False, "axes.spines.right": False})
INK, SUB, ACC, BAD, GOOD = "#1f2937", "#6b7280", "#1d4ed8", "#c2410c", "#0f766e"
GRISES = ["#6b7280", "#b9bfc9"]
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
COLS = {"Id", "Programa", "Tipo_estudiante", "Devolucion", "Fecha_pago", "Ciudad", "Fecha_nacimento", "Rango"}
COLS_ORDEN = ["Id", "Nombres", "Apellidos", "Cod_programa", "Programa", "Tipo_doc", "No_doc", "Fecha_nacimento", "Nivel", "Situacion", "Cod_estado", "Estado",
              "Cod_tipo", "Tipo_estudiante", "Pga_acomulado", "Ciudad", "Direccion", "Telefono", "Celular", "Email", "Atributo_discriminador", "Centro_costo_iceberg",
              "Rango", "Estado_retencion", "Fecha_pago", "Periodo_iceberg", "Devolucion", "Fecha_activacion"]
UMBRAL = 0.15  # error maximo tolerado para declarar confiable la proyeccion
TIPOS = ["CONTINUO", "REINGRESO", "REINICIO", "TRANSFERENCIA INTERNA", "PRIMERA VEZ"]
AQUI = Path(__file__).resolve().parent


# ---------------------------------------------------------------- carga
def nombre_seguro(t):
    return re.sub(r"[^\w-]+", "_", t)

def fecha_es(ts):
    return f"{ts.day} {MESES[ts.month - 1]} {ts.year}"

def archivos(carpeta):
    out = {}
    for ext in ("csv", "xls"):  # si hay ambos, gana el .xls (tiene encabezado)
        for f in Path(carpeta).glob(f"BAN*Levantada*.{ext}"):
            m = re.search(r"(?<!\d)(\d{6})\." + ext + "$", f.name)
            if m:
                out[m.group(1)] = f
    return dict(sorted(out.items()))

def cargar(f):
    if Path(f).suffix.lower() == ".csv":
        d = pd.read_csv(f, encoding="latin-1", header=None, names=COLS_ORDEN, dtype=str, usecols=lambda c: c in COLS)
    else:
        d = pd.read_csv(f, encoding="latin-1", sep="\t", dtype=str, usecols=lambda c: c in COLS)
    d["Fecha_pago"] = pd.to_datetime(d["Fecha_pago"], dayfirst=True, errors="coerce")
    d["Fecha_nacimento"] = pd.to_datetime(d["Fecha_nacimento"], dayfirst=True, errors="coerce")
    for c in ("Programa", "Tipo_estudiante", "Ciudad"):
        d[c] = d[c].str.replace("\xa0", " ").str.strip()
    return d

def depurar(d, bruto):
    if bruto:
        return d
    d = d[d["Devolucion"] != "Si"]
    return d.sort_values("Fecha_pago").drop_duplicates(["Id", "Programa"], keep="last")

def ancla(periodo):
    y, s = int(periodo[:4]), periodo[4:]
    return pd.Timestamp(y - 1, 11, 1) if s == "10" else pd.Timestamp(y, 5, 1)

def curva(nv, periodo):
    """Nuevos acumulados por dia respecto al ancla del ciclo."""
    t = (nv["Fecha_pago"].dt.normalize() - ancla(periodo)).dt.days.dropna().astype(int)
    if t.empty:
        return pd.Series(dtype=int)
    cnt = t.value_counts().reindex(range(t.min(), t.max() + 1), fill_value=0).sort_index()
    return cnt.cumsum()

def ticks_mes(periodo, lo, hi):
    a = ancla(periodo); ini = a + pd.Timedelta(days=int(lo))
    m = pd.Timestamp(ini.year, ini.month, 1)
    if m < ini:
        m += pd.DateOffset(months=1)
    t, l = [], []
    while (m - a).days <= hi:
        t.append((m - a).days); l.append(f"1 {MESES[m.month - 1]}"); m += pd.DateOffset(months=1)
    return t, l

def meta_para(metas, periodo):
    anio, sem = int(periodo[:4]), int(periodo[4:])
    m = metas[(metas["anio"] == anio) & (metas["sem"] == sem)]
    if m.empty:
        ult = metas[metas["sem"] == sem]["anio"].max()
        print(f"  AVISO: no hay metas {anio}-{sem}; se usan las de {ult} solo como referencia")
        m = metas[(metas["anio"] == ult) & (metas["sem"] == sem)]
        anio = ult
    return m.set_index("programa"), int(anio)


def colores_cohortes(per):
    """Cohortes previas en una rampa de grises (mas antigua = mas clara); la ultima en azul."""
    claro, oscuro = np.array(to_rgb("#b3bac5")), np.array(to_rgb("#555e6d"))
    prev = per[:-1]; col = {}
    for i, p in enumerate(prev):
        t = i / max(len(prev) - 1, 1); col[p] = to_hex(claro + (oscuro - claro) * t)
    col[per[-1]] = ACC
    return col


def _extremo(h0, h1, d0, d1):
    m = ((2 * h0 + h1) * d0 - h0 * d1) / (h0 + h1)
    if np.sign(m) != np.sign(d0):
        return 0.0
    if np.sign(d0) != np.sign(d1) and abs(m) > 3 * abs(d0):
        return 3 * d0
    return m

def _pchip(x, y, xi):
    # interpolacion cubica monotona (Fritsch-Carlson): pasa por los puntos y no crea oscilaciones
    x = np.asarray(x, float); y = np.asarray(y, float); h = np.diff(x); d = np.diff(y) / h; m = np.zeros_like(y)
    if len(x) == 2:
        m[:] = d[0]
    else:
        w1, w2 = 2 * h[1:] + h[:-1], h[1:] + 2 * h[:-1]
        dp, dn = np.where(d[:-1] == 0, 1, d[:-1]), np.where(d[1:] == 0, 1, d[1:])
        m[1:-1] = np.where(d[:-1] * d[1:] > 0, (w1 + w2) / (w1 / dp + w2 / dn), 0.0)
        m[0] = _extremo(h[0], h[1], d[0], d[1]); m[-1] = _extremo(h[-1], h[-2], d[-1], d[-2])
    i = np.clip(np.searchsorted(x, xi, side="right") - 1, 0, len(x) - 2); t = (xi - x[i]) / h[i]
    return (2 * t**3 - 3 * t**2 + 1) * y[i] + (t**3 - 2 * t**2 + t) * h[i] * m[i] + (-2 * t**3 + 3 * t**2) * y[i + 1] + (t**3 - t**2) * h[i] * m[i + 1]

def suave(s, paso=7, n=260):
    # acumulado muestreado cada `paso` dias e interpolado de forma monotona; el ultimo punto es exacto
    if len(s) < 4:
        return s.index.values, s.values
    x0, x1 = int(s.index[0]), int(s.index[-1])
    xs = sorted({x0, x1, *range(int(np.ceil(x0 / paso) * paso), x1, paso)})
    if len(xs) < 3:
        return s.index.values, s.values
    xi = np.linspace(x0, x1, n)
    return xi, _pchip(xs, s.loc[xs].values, xi)

def acum_en(s, d):
    if len(s) == 0 or d < s.index[0]:
        return 0
    return int(s.loc[min(d, s.index[-1])])

def proyectar(actual_d, refs, d):
    # cierre estimado = acumulado actual / fraccion del cierre que cada cohorte cerrada llevaba el mismo dia
    est = []
    for r in refs:
        a = acum_en(r, d)
        if a > 0 and r.iloc[-1] > 0:
            est.append(actual_d / (a / r.iloc[-1]))
    return (min(est), float(np.median(est)), max(est)) if est else None

def backtest(cur):
    # cada cohorte cerrada se proyecta con las demas y se compara con su cierre real
    ps = list(cur)
    if len(ps) < 3:
        return None
    dias = np.arange(min(s.index[0] for s in cur.values()), min(s.index[-1] for s in cur.values()) + 1); err = {}
    for c in ps:
        refs = [cur[q] for q in ps if q != c]; fin = cur[c].iloc[-1]; e = []
        for d in dias:
            r = proyectar(acum_en(cur[c], d), refs, d)
            e.append(np.nan if r is None else (r[1] - fin) / fin)
        err[c] = e
    E = pd.DataFrame(err, index=dias)
    return E, E.abs().max(axis=1)

def dia_confiable(mx, umbral=UMBRAL):
    ok = (mx <= umbral).values
    for i in range(len(ok)):
        if ok[i:].all():
            return int(mx.index[i])
    return None

def fmt_err(e):
    return "sin estimación" if not np.isfinite(e) else f"{100 * e:.0f}%"

def trozos(progs, cap=6):
    k = int(np.ceil(len(progs) / cap)); base, extra = divmod(len(progs), k); out, i = [], 0
    for j in range(k):
        n = base + (1 if j < extra else 0); out.append(progs[i:i + n]); i += n
    return out


# ---------------------------------------------------------------- graficos
def fig_curvas(titulo, subtitulo, curvas, actual, meta, xlim, ticks, ruta, meta_txt=None, colores=None, figsize=(8, 4.7), titulos=True):
    fig, ax = plt.subplots(figsize=figsize, dpi=200)
    lo, hi = xlim; span = hi - lo; etiquetas = []; ymax = meta
    colores = colores or colores_cohortes(list(curvas))
    previos = [p for p in curvas if p != actual]
    for k, p in enumerate(previos):
        s = curvas[p]
        if len(s):
            c = colores[p]; ax.plot(*suave(s), color=c, lw=1.8)
            etiquetas.append((s.index[-1], s.iloc[-1], f"{p}  {s.iloc[-1]}", c, "normal")); ymax = max(ymax, s.max())
    s = curvas[actual]
    if len(s):
        ax.plot(*suave(s), color=ACC, lw=3.0); ax.scatter([s.index[-1]], [s.iloc[-1]], s=16, color=ACC, zorder=5)
        etiquetas.append((s.index[-1], s.iloc[-1], f"{actual}  {s.iloc[-1]}", ACC, "bold")); ymax = max(ymax, s.max())
    else:
        etiquetas.append((lo, 0, f"{actual}  0", ACC, "bold"))
    ax.hlines(meta, lo, hi, color=BAD, lw=1.2, ls="--")
    ax.text(lo + span * 0.01, meta + ymax * 0.015, meta_txt or f"Meta {meta}", color=BAD, fontsize=9, va="bottom")
    etiquetas.sort(key=lambda e: e[1]); sep = ymax * 0.06; ys = []
    for e in etiquetas:
        ys.append(e[1] if not ys else max(e[1], ys[-1] + sep))
    for e, y in zip(etiquetas, ys):
        ax.text(e[0] + span * 0.012, y, e[2], color=e[3], va="center", fontsize=8.5, fontweight=e[4],
                bbox=dict(facecolor="white", edgecolor="none", pad=1.2, alpha=.85))
    ax.set_xlim(lo, hi + span * 0.24); ax.set_ylim(0, ymax * 1.14)
    ax.set_xticks(ticks[0]); ax.set_xticklabels(ticks[1], fontsize=9)
    ax.yaxis.get_major_locator().set_params(integer=True)
    ax.grid(axis="y", alpha=.2); ax.set_ylabel("Nuevos acumulados", fontsize=9)
    if titulos:
        ax.set_title(titulo, loc="left", fontsize=12, fontweight="bold", color=INK, pad=24)
        ax.text(0, 1.03, subtitulo, transform=ax.transAxes, fontsize=8.3, color=SUB)
    fig.tight_layout(); fig.savefig(ruta); plt.close(fig)

def fig_grid(progs, curs, metas_p, per, actual, xlim, ticks, colores, ruta):
    # una lamina con pocas curvas grandes (2x2 o 3x2), curvas semanales suavizadas
    n = len(progs); ncols = 2 if n <= 4 else 3; nrows = int(np.ceil(n / ncols)); lo, hi = xlim
    fig, axs = plt.subplots(nrows, ncols, figsize=(12.3, 5.35), dpi=200, squeeze=False)
    for ax, prog in zip(axs.ravel(), progs):
        m = int(metas_p[prog]); ymax = m
        for p in per:
            s = curs[prog][p]
            if len(s):
                ax.plot(*suave(s), color=colores[p], lw=2.8 if p == actual else 1.6, zorder=3 if p == actual else 2); ymax = max(ymax, s.max())
        sa = curs[prog][actual]
        if len(sa):
            ax.scatter([sa.index[-1]], [sa.iloc[-1]], s=16, color=ACC, zorder=5)
        ax.axhline(m, color=BAD, lw=1.0, ls="--")
        na = int(sa.iloc[-1]) if len(sa) else 0
        ax.set_title(f"{prog}   {na} / {m}", fontsize=11, loc="left", color=INK, fontweight="bold")
        ax.set_xlim(lo, hi); ax.set_ylim(0, ymax * 1.12)
        ax.set_xticks(ticks[0]); ax.set_xticklabels(ticks[1], fontsize=8)
        ax.tick_params(axis="y", labelsize=8); ax.yaxis.get_major_locator().set_params(integer=True); ax.grid(axis="y", alpha=.2)
    for ax in axs.ravel()[n:]:
        ax.axis("off")
    hs = [Line2D([0], [0], color=colores[p], lw=2.8 if p == actual else 1.6, label=p) for p in per] + [Line2D([0], [0], color=BAD, lw=1, ls="--", label="Meta")]
    fig.legend(handles=hs, loc="lower center", ncol=len(hs), frameon=False, fontsize=9.5)
    fig.tight_layout(rect=(0, .07, 1, 1)); fig.savefig(ruta); plt.close(fig)

def fig_proyeccion(tot, actual, d, est, confiable, meta, xlim, ticks, colores, ruta):
    fig, ax = plt.subplots(figsize=(8.4, 5.1), dpi=200); lo, hi = xlim; span = hi - lo; ymax = meta; etiquetas = []
    for p, s in tot.items():
        if not len(s):
            continue
        ax.plot(*suave(s), color=colores[p], lw=2.8 if p == actual else 1.6, zorder=3 if p == actual else 2); ymax = max(ymax, s.max())
        if p != actual:
            ax.scatter([d], [acum_en(s, d)], s=16, color=colores[p], zorder=4)
            etiquetas.append((s.iloc[-1], f"{p}  {s.iloc[-1]}", colores[p]))
    ca = acum_en(tot[actual], d); ax.scatter([d], [ca], s=34, color=ACC, zorder=5)
    ax.axvline(d, color=INK, lw=.8, ls=":"); ax.text(d, ymax * 1.07, f"corte: {ca}", ha="center", fontsize=8.5, color=INK)
    ax.hlines(meta, lo, hi, color=BAD, lw=1.1, ls="--"); ax.text(lo + span * .01, meta + ymax * .015, f"Meta {meta}", color=BAD, fontsize=9)
    if est and confiable:
        l, m, h = est
        ax.vlines(hi, l, h, color=ACC, lw=8, alpha=.28); ax.plot([d, hi], [ca, m], color=ACC, lw=1.5, ls=(0, (4, 3)))
        etiquetas.append((m, f"proyección {m:.0f}", ACC)); ymax = max(ymax, h)
    etiquetas.sort(key=lambda e: e[0]); sep = ymax * .06; ys = []
    for e in etiquetas:
        ys.append(e[0] if not ys else max(e[0], ys[-1] + sep))
    for e, y in zip(etiquetas, ys):
        ax.text(hi + span * .02, y, e[1], color=e[2], va="center", fontsize=8.5, fontweight="bold" if e[2] == ACC else "normal")
    ax.set_xlim(lo, hi + span * .27); ax.set_ylim(0, ymax * 1.14)
    ax.set_xticks(ticks[0]); ax.set_xticklabels(ticks[1], fontsize=9); ax.yaxis.get_major_locator().set_params(integer=True)
    ax.grid(axis="y", alpha=.2); ax.set_ylabel("Nuevos acumulados", fontsize=9)
    fig.tight_layout(); fig.savefig(ruta); plt.close(fig)

def fig_backtest(bts, ruta):
    fig, axs = plt.subplots(len(bts), 1, figsize=(7.6, 2.85 * len(bts) + .3), dpi=200, squeeze=False)
    cols = ["#374151", "#6b7280", "#a5acb8"]
    for ax, b in zip(axs.ravel(), bts):
        E = 100 * b["E"]; lo, hi = E.index[0], E.index[-1]
        ax.axhspan(-100 * UMBRAL, 100 * UMBRAL, color=GOOD, alpha=.13)
        for k, c in enumerate(E.columns):
            ax.plot(E.index, E[c].clip(-60, 60), lw=1.6, color=cols[k % 3], label=c)
        ax.axhline(0, color=INK, lw=.8)
        if b["dstar"] is not None:
            ax.axvline(b["dstar"], color=ACC, lw=1.4)
            ax.text(b["dstar"] + 1, 49, f"error máximo ≤ {100 * UMBRAL:.0f}% desde {fecha_es(ancla(b['ref']) + pd.Timedelta(days=b['dstar']))}", color=ACC, fontsize=8.5, fontweight="bold")
        if b["d_hoy"] is not None and lo <= b["d_hoy"] <= hi:
            ax.axvline(b["d_hoy"], color=BAD, lw=1.2, ls=":"); ax.text(b["d_hoy"] + 1, -55, "corte actual", color=BAD, fontsize=8.5, bbox=dict(facecolor="white", edgecolor="none", pad=1.2, alpha=.9))
        tk = ticks_mes(b["ref"], lo, hi); ax.set_xticks(tk[0]); ax.set_xticklabels(tk[1], fontsize=8.5)
        ax.set_xlim(lo, hi); ax.set_ylim(-60, 60); ax.tick_params(axis="y", labelsize=8.5); ax.grid(axis="y", alpha=.2)
        ax.set_ylabel("Error de la proyección (%)", fontsize=8.5)
        ax.set_title(f"Semestre {b['sem'][0]}: prueba retrospectiva", loc="left", fontsize=9.5, fontweight="bold", color=INK, pad=14)
        ax.legend(fontsize=8, frameon=False, loc="lower right", bbox_to_anchor=(1.0, 1.0), ncol=len(E.columns))
    fig.text(.01, .004, "Errores mayores a 60% recortados en el gráfico.", fontsize=7.5, color=SUB)
    fig.tight_layout(rect=(0, .02, 1, 1)); fig.savefig(ruta); plt.close(fig)

def barras_brecha(ax, g, titulo):
    cols = [BAD if b < 0 else GOOD for b in g.brecha]
    ax.barh(g.index, g.brecha, color=cols, height=.68)
    lim = max(abs(g.brecha).max(), 5) * 1.9
    for i, (p, r) in enumerate(g.iterrows()):
        ax.text(r.brecha + (lim * .02 if r.brecha >= 0 else -lim * .02), i, f"{int(r.n)}/{int(r.meta)}  ({int(r.brecha):+d})",
                va="center", ha="left" if r.brecha >= 0 else "right", fontsize=8, color=INK)
    ax.axvline(0, color=INK, lw=1); ax.set_xlim(-lim, lim); ax.tick_params(axis="y", length=0)
    ax.grid(axis="x", alpha=.2); ax.set_title(titulo, loc="left", fontsize=11, fontweight="bold", color=INK)
    ax.set_xlabel("Nuevos por encima (+) o por debajo (-) de la meta", fontsize=9)

def resumen_brecha(g):
    neg, pos = int(g.brecha[g.brecha < 0].sum()), int(g.brecha[g.brecha > 0].sum())
    return f"{(g.brecha < 0).sum()} programas suman {neg}, {(g.brecha > 0).sum()} suman +{pos}, neto {neg + pos}"

def fig_brecha(periodo, g, corte, ruta):
    fig, ax = plt.subplots(figsize=(9, 7), dpi=200)
    barras_brecha(ax, g, f"Brecha frente a la meta, cohorte {periodo}: {resumen_brecha(g)}")
    fig.text(.01, .008, f"Primera vez, sin devoluciones ni duplicados. Datos al {fecha_es(corte)}.", fontsize=8, color=SUB)
    fig.tight_layout(rect=(0, .02, 1, 1)); fig.savefig(ruta); plt.close(fig)

def fig_comparativo(anio, g10, g20, ruta):
    orden = g10.sort_values("brecha").index
    fig, axs = plt.subplots(1, 2, figsize=(14, 7.2), dpi=200, sharey=True)
    barras_brecha(axs[0], g10.loc[orden], f"{anio}-1: {resumen_brecha(g10)}")
    barras_brecha(axs[1], g20.loc[orden], f"{anio}-2: {resumen_brecha(g20)}")
    axs[1].tick_params(axis="y", length=0)
    fig.tight_layout(); fig.savefig(ruta); plt.close(fig)

def fig_persistencia(filas, ruta, titulos=True):
    k = max(len(f[2]) for f in filas)
    M = np.full((len(filas), k), np.nan)
    for i, f in enumerate(filas):
        M[i, :len(f[2])] = f[2]
    fig, ax = plt.subplots(figsize=(7.5, 0.6 * len(filas) + 2.2), dpi=200)
    ax.imshow(np.ma.masked_invalid(M), cmap="Blues", vmin=.55, vmax=1.0, aspect="auto")
    for i in range(M.shape[0]):
        for j in range(M.shape[1]):
            if not np.isnan(M[i, j]):
                ax.text(j, i, f"{100 * M[i, j]:.0f}%", ha="center", va="center", fontsize=9, color="white" if M[i, j] > .82 else INK)
    ax.set_yticks(range(len(filas))); ax.set_yticklabels([f"{f[0]}  (n={f[1]})" for f in filas], fontsize=9)
    ax.set_xticks(range(k)); ax.set_xticklabels([f"+{j + 1}" for j in range(k)], fontsize=9)
    ax.set_xlabel("Semestres después de la matrícula de ingreso", fontsize=9)
    for s in ax.spines.values():
        s.set_visible(False)
    if titulos:
        ax.set_title("Permanencia de los nuevos: % que vuelve a aparecer pagado", loc="left", fontsize=12, fontweight="bold", color=INK, pad=22)
        ax.text(0, 1.02, "Reaparecer no distingue deserción de grado, pausa o cambio de programa.", transform=ax.transAxes, fontsize=8.3, color=SUB)
    fig.tight_layout(); fig.savefig(ruta); plt.close(fig)


# ---------------------------------------------------------------- tablas
def tabla_cumplimiento(actual, NV, meta, previos):
    n = NV[actual].Programa.value_counts().reindex(meta.index).fillna(0).astype(int)
    t = pd.DataFrame({"Escuela": meta.escuela, "Meta": meta.meta, "Nuevos": n})
    t["Cumplimiento_%"] = (100 * t.Nuevos / t.Meta).round(1)
    t["Brecha"] = t.Nuevos - t.Meta
    t["LogChange_%"] = np.where(t.Nuevos > 0, 100 * np.log(t.Nuevos.clip(lower=1) / t.Meta), np.nan).round(1)
    t["SPC_%"] = (200 * (t.Nuevos - t.Meta) / (t.Nuevos + t.Meta)).round(1)
    for p in previos:
        t[f"Nuevos_{p}"] = NV[p].Programa.value_counts().reindex(meta.index).fillna(0).astype(int)
    t = t.sort_values("Brecha")
    tot = {c: t[c].sum() for c in t.columns if c.startswith("Nuevos") or c in ("Meta", "Brecha")}
    tot["Cumplimiento_%"] = round(100 * t.Nuevos.sum() / t.Meta.sum(), 1)
    t.loc["TOTAL"] = pd.Series(tot)
    return t.rename_axis("Programa")

def tabla_comparativa(p, DEP, meta, previo=None):
    """Tabla estilo notebook: tipos de estudiante, meta, cumplimiento y cierre del mismo semestre del anio anterior."""
    dep = DEP[p]; c = pd.crosstab(dep.Programa, dep.Tipo_estudiante).reindex(meta.index).fillna(0).astype(int)
    t = pd.DataFrame({"Escuela": meta.escuela})
    pv = c["PRIMERA VEZ"] if "PRIMERA VEZ" in c else 0
    t["Primera vez"] = pv; t["Meta"] = meta.meta
    t["Cumpl. %"] = 100 * t["Primera vez"] / t.Meta
    t["Dif. meta"] = t["Primera vez"] - t.Meta
    t["SPC %"] = 200 * (t["Primera vez"] - t.Meta) / (t["Primera vez"] + t.Meta).replace(0, np.nan)
    if previo is not None:
        pp, vp = previo
        t[f"Cierre {pp}"] = vp.reindex(meta.index).fillna(0).astype(int)
    nombres = {"CONTINUO": "Continuo", "REINGRESO": "Reingreso", "REINICIO": "Reinicio", "TRANSFERENCIA INTERNA": "Trans. interna"}
    for k, v in nombres.items():
        t[v] = c[k] if k in c else 0
    t["Total"] = c.sum(axis=1)
    return t.sort_values("Escuela", kind="stable").rename_axis("Programa")

def resumen_escuelas(t):
    num = [x for x in t.columns if x not in ("Escuela", "Cumpl. %", "SPC %")]
    e = t.groupby("Escuela")[num].sum()
    e["Cumpl. %"] = 100 * e["Primera vez"] / e.Meta
    e["SPC %"] = 200 * (e["Primera vez"] - e.Meta) / (e["Primera vez"] + e.Meta)
    e.loc["TOTAL"] = e.sum(); e.loc["TOTAL", "Cumpl. %"] = 100 * e.loc["TOTAL", "Primera vez"] / e.loc["TOTAL", "Meta"]
    e.loc["TOTAL", "SPC %"] = 200 * (e.loc["TOTAL", "Primera vez"] - e.loc["TOTAL", "Meta"]) / (e.loc["TOTAL", "Primera vez"] + e.loc["TOTAL", "Meta"])
    cols = [c for c in t.columns if c in e.columns]
    return e[cols].rename_axis("Escuela")

def formato_tabla(t):
    """Texto listo para diapositiva: enteros sin decimales, porcentajes con uno, signo en diferencias."""
    o = t.copy().astype(object)
    for c in o.columns:
        if c == "Escuela":
            continue
        if c in ("Cumpl. %", "SPC %"):
            o[c] = t[c].map(lambda v: "" if pd.isna(v) else f"{v:.1f}")
        elif c == "Dif. meta":
            o[c] = t[c].map(lambda v: f"{v:+,.0f}")
        else:
            o[c] = t[c].map(lambda v: f"{v:,.0f}")
    return o

def tabla_tipos(dep, meta):
    c = pd.crosstab(dep.Programa, dep.Tipo_estudiante)
    c = c.reindex(meta.index).fillna(0).astype(int)
    cols = [t for t in TIPOS if t in c.columns] + [t for t in c.columns if t not in TIPOS]
    c = c[cols]; c["TOTAL"] = c.sum(axis=1)
    c.loc["TOTAL"] = c.sum()
    return c.rename_axis("Programa")

def guardar_excel(hojas, ruta):
    try:
        import openpyxl  # noqa
    except ImportError:
        carpeta = ruta.parent / "tablas"; carpeta.mkdir(exist_ok=True)
        for k, v in hojas.items():
            v.to_csv(carpeta / f"{k}.csv", encoding="utf-8-sig")
        print("  openpyxl no instalado: tablas guardadas como CSV en", carpeta); return
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        for k, v in hojas.items():
            v.to_excel(w, sheet_name=k[:31]); ws = w.sheets[k[:31]]
            for col in ws.columns:
                ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(ancho + 2, 42)
            ws.freeze_panes = "B2"


def a_igual_dia(nv, periodo, d):
    t = (nv["Fecha_pago"].dt.normalize() - ancla(periodo)).dt.days
    return t.le(d).groupby(nv["Programa"]).sum()

# ---------------------------------------------------------------- documento con diapositivas
def generar_deck(ruta, laminas, pie):
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        from PIL import Image
    except ImportError:
        print("  Falta python-pptx: pip install python-pptx   (se omite el documento)"); return False
    rgb = lambda h: RGBColor.from_string(h.lstrip("#").upper())
    prs = Presentation(); prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    def txt(sl, x, y, w, h, parrafos, size=14, bold=False, color=INK, align=None):
        tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame; tf.word_wrap = True
        for i, t in enumerate([parrafos] if isinstance(parrafos, str) else parrafos):
            o = {}
            if isinstance(t, tuple):
                t, o = t
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            r = p.add_run(); r.text = t; r.font.size = Pt(o.get("size", size)); r.font.bold = o.get("bold", bold); r.font.color.rgb = rgb(o.get("color", color)); r.font.name = "Calibri"
            p.space_after = Pt(6)
            if align:
                p.alignment = align
        return tb

    def base(titulo, sub=None):
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        txt(sl, 0.5, 0.3, 12.3, 0.8, titulo, 26, True)
        if sub:
            txt(sl, 0.5, 1.02, 12.3, 0.4, sub, 13, color=SUB)
        txt(sl, 0.5, 7.08, 12.3, 0.3, pie, 9, color=SUB)
        return sl

    def imagen(sl, ruta_png, x=0.5, y=1.5, w=12.3, h=5.5):
        iw, ih = Image.open(ruta_png).size; r = min(w / iw, h / ih)
        sl.shapes.add_picture(str(ruta_png), Inches(x + (w - iw * r) / 2), Inches(y), Inches(iw * r), Inches(ih * r))

    def celda(c, t, size, bold=False, color=INK, fill=None, align=None):
        c.text = ""; p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = str(t); r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = rgb(color); r.font.name = "Calibri"
        c.margin_top = c.margin_bottom = Inches(0.015); c.margin_left = c.margin_right = Inches(0.06)
        if align:
            p.alignment = align
        if fill:
            c.fill.solid(); c.fill.fore_color.rgb = rgb(fill)

    for L in laminas:
        tipo = L[0]
        if tipo == "portada":
            sl = prs.slides.add_slide(prs.slide_layouts[6])
            barra = sl.shapes.add_shape(1, Inches(0.5), Inches(2.55), Inches(0.12), Inches(1.5)); barra.fill.solid(); barra.fill.fore_color.rgb = rgb(ACC); barra.line.fill.background()
            txt(sl, 0.85, 2.4, 11.5, 1.0, L[1], 38, True); txt(sl, 0.85, 3.35, 11.5, 1.6, L[2], 16, color=SUB)
            txt(sl, 0.85, 6.6, 11.5, 0.5, pie, 10, color=SUB)
        elif tipo == "resumen":
            sl = base(L[1], L[2]); kp = L[3]; ancho = (12.3 - 0.25 * (len(kp) - 1)) / max(len(kp), 1)
            for i, k in enumerate(kp):
                x = 0.5 + i * (ancho + 0.25)
                f = sl.shapes.add_shape(1, Inches(x), Inches(1.7), Inches(ancho), Inches(2.25)); f.fill.solid(); f.fill.fore_color.rgb = rgb("#f3f4f6"); f.line.fill.background()
                txt(sl, x + 0.15, 1.78, ancho - 0.3, 0.4, k["label"], 12, True, SUB)
                txt(sl, x + 0.15, 2.15, ancho - 0.3, 0.7, k["valor"], 32, True, ACC)
                txt(sl, x + 0.15, 2.85, ancho - 0.3, 1.1, k["sub"], 11, color=INK)
            txt(sl, 0.5, 4.25, 12.3, 2.7, ["•  " + n for n in L[4]], 14)
        elif tipo == "imagen":
            sl = base(L[1], L[2]); imagen(sl, L[3])
        elif tipo == "texto":
            sl = base(L[1], L[2]); txt(sl, 0.5, 1.6, 12.3, 5.3, L[3], 15)
        elif tipo == "mixto":
            sl = base(L[1], L[2]); rimg, par, lado, frac = L[3], L[4], L[5], L[6]; wi = 12.3 * frac; wt = 12.3 - wi - 0.3
            if lado == "izq":
                imagen(sl, rimg, 0.5, 1.5, wi, 5.45); txt(sl, 0.5 + wi + 0.3, 1.55, wt, 5.3, par, 12)
            else:
                txt(sl, 0.5, 1.55, wt, 5.3, par, 12); imagen(sl, rimg, 0.5 + wt + 0.3, 1.5, wi, 5.45)
        elif tipo == "tabla":
            sl = base(L[1], L[2]); df = L[3]; filas, cols = len(df) + 1, len(df.columns) + 1
            alto = min(0.23, 5.3 / filas); ancho_t = L[4] if len(L) > 4 else 11.0
            tb = sl.shapes.add_table(filas, cols, Inches(0.5), Inches(1.55), Inches(ancho_t), Inches(alto * filas)).table
            tb.columns[0].width = Inches(ancho_t * 0.30)
            for j in range(1, cols):
                tb.columns[j].width = Inches(ancho_t * 0.70 / (cols - 1))
            tam = 9 if filas > 14 else 11
            celda(tb.cell(0, 0), df.index.name or "", tam, True, "#ffffff", INK)
            for j, c in enumerate(df.columns, 1):
                celda(tb.cell(0, j), c, tam, True, "#ffffff", INK, PP_ALIGN.RIGHT)
            for i, (ix, fila) in enumerate(df.iterrows(), 1):
                tot = str(ix) == "TOTAL"; fill = "#e5e7eb" if tot else ("#f9fafb" if i % 2 == 0 else "#ffffff")
                celda(tb.cell(i, 0), ix, tam, tot, INK, fill)
                for j, v in enumerate(fila, 1):
                    celda(tb.cell(i, j), "" if pd.isna(v) else (f"{v:,.0f}" if isinstance(v, (int, float, np.integer, np.floating)) else v), tam, tot, INK, fill, PP_ALIGN.RIGHT)
            for r in tb.rows:
                r.height = Inches(alto)
            if len(L) > 5 and L[5]:
                txt(sl, 0.5 + ancho_t + 0.3, 1.55, 12.83 - (0.5 + ancho_t + 0.3), 5, L[5], 12)
    prs.save(ruta); return True


# ---------------------------------------------------------------- principal
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--datos", default=str(AQUI)); ap.add_argument("--salida", default=str(AQUI / "Salida"))
    ap.add_argument("--metas", default=str(AQUI / "metas.csv"))
    ap.add_argument("--anios", type=int, default=0, help="cohortes a mostrar por semestre (0 = todas)")
    ap.add_argument("--bruto", action="store_true", help="no quitar devoluciones ni duplicados")
    ap.add_argument("--cerrado", action="store_true", help="tratar el ultimo periodo como cerrado (no en curso)")
    ap.add_argument("--figuras", action="store_true", help="guardar tambien las figuras sueltas (incluye una por programa)")
    ap.add_argument("--fecha-reporte", default=None, help="fecha del reporte AAAA-MM-DD (por defecto, hoy)")
    ap.add_argument("--por-lamina", type=int, default=2, help="programas por diapositiva en cada escuela (por defecto 2)")
    a = ap.parse_args()
    rep_fecha = pd.Timestamp(a.fecha_reporte).normalize() if a.fecha_reporte else pd.Timestamp.today().normalize()
    out = Path(a.salida); out.mkdir(exist_ok=True)
    figdir = (out / "Figuras") if a.figuras else Path(tempfile.mkdtemp())
    figdir.mkdir(exist_ok=True)
    files = archivos(a.datos)
    if not files:
        sys.exit("No encontre reportes BAN en " + a.datos)
    metas = pd.read_csv(a.metas, comment="#")
    D = {p: cargar(f) for p, f in files.items()}
    DEP = {p: depurar(d, a.bruto) for p, d in D.items()}
    NV = {p: d[d.Tipo_estudiante == "PRIMERA VEZ"] for p, d in DEP.items()}
    corte = {p: D[p].Fecha_pago.max() for p in D}
    ultimo = list(files)[-1]
    en_curso = {ultimo} if (not a.cerrado and (rep_fecha - corte[ultimo]).days <= 30) else set()
    print("Periodos:", ", ".join(files), "|", "bruto" if a.bruto else "depurado", "| en curso:", ", ".join(en_curso) or "ninguno")
    todos_prog = set(metas.programa); hojas, lineas, brechas = {}, [], {}
    lam_sem, kpis, notas, seg_lamina = [], [], [], None
    lam_proy, bts = [], []

    cal = []
    for p, d in D.items():
        pv = d[d.Tipo_estudiante == "PRIMERA VEZ"]; nv = NV[p]
        cal.append({"Periodo": p, "En_curso": p in en_curso, "Filas": len(d), "PV_bruto": len(pv), "PV_devoluciones": int((pv.Devolucion == "Si").sum()),
                    "PV_Id_duplicados": int(pv.duplicated(["Id", "Programa"]).sum()), "PV_usado": len(nv),
                    "PV_fuera_de_metas": int((~nv.Programa.isin(todos_prog)).sum()), "Ultimo_pago": corte[p].date()})
    hojas["Calidad_datos"] = pd.DataFrame(cal).set_index("Periodo")
    hojas["Metas_usadas"] = metas.set_index("anio")

    for sem in ("10", "20"):
        per = [p for p in files if p.endswith(sem)]
        if a.anios:
            per = per[-a.anios:]
        if not per:
            continue
        actual = per[-1]; previos = per[:-1]; cerr = [p for p in per if p not in en_curso]
        meta, anio_meta = meta_para(metas, actual); ref = anio_meta != int(actual[:4]); encurso = actual in en_curso
        print(f"Cohorte {actual}{' (en curso)' if encurso else ''}: {len(per)} cohortes, {len(meta)} programas")
        nvm = {p: NV[p][NV[p].Programa.isin(meta.index)] for p in per}
        tot = {p: curva(nvm[p], p) for p in per}
        lo = min(s.index[np.argmax(s.values >= max(1, .02 * s.iloc[-1]))] for s in tot.values() if len(s)) - 7
        hi = max(s.index[-1] for s in tot.values() if len(s)); tk = ticks_mes(actual, lo, hi); col = colores_cohortes(per)
        sub = f"Cohorte {actual} frente a {', '.join(previos) or 'sin cohortes previas'}  ·  primera vez, " + \
              ("conteo bruto" if a.bruto else "sin devoluciones ni duplicados") + f"  ·  datos al {fecha_es(corte[actual])}  ·  curvas semanales suavizadas"
        mtxt = lambda m: f"Meta {anio_meta}-{sem[0]} (referencia): {m}" if ref else f"Meta {m}"

        def titulo(nombre, n, m):
            return f"{nombre}: {n} nuevos a la fecha" + ("" if ref else f", meta {m}") if (ref or encurso) else f"{nombre}: {n} de {m} ({100 * n / m:.1f}% de la meta)"
        curs = {prog: {p: curva(nvm[p][nvm[p].Programa == prog], p) for p in per} for prog in meta.index}
        if a.figuras:
            carpeta = figdir / f"Curvas_{sem}"; carpeta.mkdir(exist_ok=True)
            for prog in meta.index:
                cur = curs[prog]; n = int(cur[actual].iloc[-1]) if len(cur[actual]) else 0; m = int(meta.loc[prog, "meta"])
                fig_curvas(titulo(prog, n, m), sub, cur, actual, m, (lo, hi), tk, carpeta / f"Curva_{sem}_{nombre_seguro(prog)}.png", mtxt(m), col)
        n = int(tot[actual].iloc[-1]); m = int(meta.meta.sum())
        ruta_utb = figdir / f"UTB_curva_{sem}.png"
        fig_curvas(titulo("UTB", n, m), sub, tot, actual, m, (lo, hi), tk, ruta_utb, mtxt(m), col, figsize=(11.6, 5.2), titulos=False)

        # cohortes cerradas
        for p in cerr:
            if not ((metas["anio"] == int(p[:4])) & (metas["sem"] == int(p[4:]))).any():
                continue
            mp, _ = meta_para(metas, p)
            g = pd.DataFrame({"n": NV[p].Programa.value_counts().reindex(mp.index).fillna(0).astype(int), "meta": mp.meta})
            g["brecha"] = g.n - g.meta; brechas[p] = g.sort_values("brecha")
        if cerr and cerr[-1] in brechas:
            c = cerr[-1]; mc, _ = meta_para(metas, c)
            hojas[f"Cumplimiento_{c}"] = tabla_cumplimiento(c, NV, mc, [p for p in cerr if p != c])
            hojas[f"Tipos_{c}"] = tabla_tipos(DEP[c], mc)
            nc = int(brechas[c].n.sum()); mm = int(brechas[c].meta.sum()); b = brechas[c]
            lineas.append(f"{c} (cerrada): {nc} nuevos de {mm} (meta), {100 * nc / mm:.1f}%; {resumen_brecha(b)}")
            kpis.append({"label": f"{c} cerrada", "valor": f"{nc} / {mm}", "sub": f"{100 * nc / mm:.1f}% de la meta. {resumen_brecha(b)}."})
            peor = b[b.brecha < 0].head(3)
            if len(peor):
                notas.append(f"Mayores brechas en {c}: " + ", ".join(f"{p.title()} ({int(r.brecha)})" for p, r in peor.iterrows()) + ".")

        # lamina UTB del semestre
        if encurso:
            d = (corte[actual] - ancla(actual)).days
            cols_sg = {f"Nuevos_{actual}": NV[actual].Programa.value_counts()}
            for p in previos:
                cols_sg[f"{p}_mismo_dia"] = a_igual_dia(NV[p], p, d)
            sg = pd.DataFrame(cols_sg).reindex(meta.index).fillna(0).astype(int)
            sg.insert(0, "Meta_ref" if ref else "Meta", meta.meta); sg.loc["TOTAL"] = sg.sum()
            hojas[f"Seguimiento_{actual}"] = sg.rename_axis("Programa")
            tt = sg.loc["TOTAL"]; ant = previos[-1] if previos else None
            tit = f"Semestre {sem[0]}: {actual} lleva {int(tt[f'Nuevos_{actual}'])} nuevos" + (f"; a igual día, {ant} llevaba {int(tt[f'{ant}_mismo_dia'])}" if ant else "")
            lineas.append(f"{actual} EN CURSO (dia +{d}, datos al {fecha_es(corte[actual])}): {int(tt[f'Nuevos_{actual}'])} nuevos; a igual dia " +
                          ", ".join(f"{p}: {int(tt[f'{p}_mismo_dia'])}" for p in previos) +
                          (f". Falta cargar la meta {actual[:4]}-{sem[0]} en metas.csv." if ref else ""))
            kpis.insert(0, {"label": f"{actual} en curso", "valor": f"{int(tt[f'Nuevos_{actual}'])}",
                            "sub": f"nuevos al {fecha_es(corte[actual])}. Meta {int(tt['Meta' if not ref else 'Meta_ref'])}. A igual día: " + ", ".join(f"{p} {int(tt[f'{p}_mismo_dia'])}" for p in previos[-3:]) + "."})
            ren = {f"Nuevos_{actual}": f"{actual} (hoy)", **{f"{p}_mismo_dia": f"{p} a igual día" for p in previos}}
            seg_lamina = ("tabla", f"Seguimiento de {actual} por programa, a igual día", f"Día +{d} del ciclo (ancla {fecha_es(ancla(actual))}). Compara ritmo, no proyecta el cierre.", sg.rename(columns=ren).rename_axis("Programa"), 9.5)
        else:
            tit = f"Semestre {sem[0]}: {actual} cerró con {n} de {m} ({100 * n / m:.1f}% de la meta)"
        lam = [("imagen", tit, sub, ruta_utb)]
        objetivos = [(actual, meta, ref, encurso)]
        if encurso and cerr and ((metas["anio"] == int(cerr[-1][:4])) & (metas["sem"] == int(cerr[-1][4:]))).any():
            objetivos.insert(0, (cerr[-1], meta_para(metas, cerr[-1])[0], False, False))
        for pc, mc_, rf_, ec_ in objetivos:
            prev_y = f"{int(pc[:4]) - 1}{sem}"
            previo = (prev_y, NV[prev_y].Programa.value_counts()) if prev_y in NV else None
            tc = tabla_comparativa(pc, DEP, mc_, previo); te = resumen_escuelas(tc)
            hojas[f"Comparativa_{pc}"] = tc; hojas[f"Comparativa_escuelas_{pc}"] = te
            nota_c = ("En curso: el cumplimiento es parcial, la cohorte aún suma pagos. " if ec_ else "Cohorte cerrada. ") + \
                     ("Metas de referencia (no hay meta cargada para este año). " if rf_ else "") + \
                     "Cumpl. % = primera vez / meta. SPC % = 200·(real − meta)/(real + meta), simétrico respecto a la meta."
            etq = " (en curso)" if ec_ else " (cerrada)"
            lam.append(("tabla", f"Comparativa por programa, {pc}{etq}", f"Datos al {fecha_es(corte[pc])}. " + nota_c, formato_tabla(tc), 12.3))
            lam.append(("tabla", f"Resumen por escuela, {pc}{etq}", f"Datos al {fecha_es(corte[pc])}. " + nota_c, formato_tabla(te), 12.3))
        hojas[f"Curva_UTB_diaria_{sem}"] = pd.DataFrame(tot).sort_index().ffill().fillna(0).astype(int).rename_axis("Dia_desde_ancla")
        cur_c = {p: tot[p] for p in cerr if len(tot[p])}; bt = backtest(cur_c)
        if bt is not None:
            E, mx = bt; mxf = mx.fillna(np.inf); dstar = dia_confiable(mxf)
            dh = (corte[actual] - ancla(actual)).days if encurso else None
            bts.append({"sem": sem, "E": E, "mx": mxf, "dstar": dstar, "d_hoy": dh, "ref": actual})
            hojas[f"Proyeccion_prueba_{sem}"] = (100 * E).round(1).rename_axis("Dia"); hojas[f"Proyeccion_prueba_{sem}"]["Error_max_abs"] = (100 * mx).round(1)
            if encurso:
                est = proyectar(acum_en(tot[actual], dh), list(cur_c.values()), dh)
                err_hoy = float(mxf.loc[min(max(dh, mxf.index[0]), mxf.index[-1])]); ok = est is not None and err_hoy <= UMBRAL
                mt = int(meta.meta.sum()); rpj = figdir / f"Proyeccion_{actual}.png"
                fig_proyeccion(tot, actual, dh, est, ok, mt, (lo, hi), tk, col, rpj)
                desde = fecha_es(ancla(actual) + pd.Timedelta(days=dstar)) if dstar is not None else None
                if est:
                    mec = f"A esta fecha el cálculo daría {est[1]:.0f} nuevos (rango {est[0]:.0f} a {est[2]:.0f})."
                else:
                    mec = "Aún no hay datos suficientes para estimar."
                par = [("Cierre proyectado", {"size": 12, "bold": True, "color": SUB}),
                       (f"{est[1]:.0f}" if ok else "No confiable aún", {"size": 30, "bold": True, "color": ACC}),
                       (f"Rango {est[0]:.0f} a {est[2]:.0f}. Meta {mt}." if ok else mec, {"size": 12}),
                       (f"Error máximo de la prueba retrospectiva a esta fecha: {fmt_err(err_hoy)}. Umbral para confiar: {100 * UMBRAL:.0f}%.", {"size": 12}),
                       (f"Según esa prueba, será confiable desde {desde}." if desde else "En las cohortes anteriores el error no bajó de forma sostenida del umbral.", {"size": 12}),
                       ("Mientras tanto, lo que sí informa es el ritmo frente a las cohortes anteriores a igual día (diapositiva de seguimiento). Los puntos de la gráfica marcan dónde estaba cada cohorte ese día.", {"size": 12, "color": SUB})]
                lam_proy.append(("mixto", f"Proyección de cierre de {actual}: " + (f"{est[1]:.0f} nuevos (rango {est[0]:.0f} a {est[2]:.0f})" if ok else "todavía no es confiable"),
                                 f"Cohorte {actual} frente a cohortes cerradas del mismo semestre  ·  datos al {fecha_es(corte[actual])}", rpj, par, "izq", 0.66))
                lineas.append(f"Proyeccion {actual}: " + (f"{est[1]:.0f} (rango {est[0]:.0f} a {est[2]:.0f})" if ok else f"no confiable aun (error de prueba a esta fecha {fmt_err(err_hoy)}); calculo mecanico {est[1]:.0f}" if est else "sin datos") + (f"; confiable desde {desde}" if desde else ""))
                notas.append(f"Proyección de cierre de {actual}: " + (f"{est[1]:.0f} nuevos (rango {est[0]:.0f} a {est[2]:.0f})." if ok else "todavía no es confiable") + (f" Según la prueba retrospectiva lo será desde {desde}." if (desde and not ok) else ""))
        escs = [e for e in ["IA&D", "NL&S", "TD"] if e in set(meta.escuela)] + sorted(set(meta.escuela) - {"IA&D", "NL&S", "TD"})
        for esc in escs:
            progs = list(meta[meta.escuela == esc].sort_values("meta", ascending=False).index); gr = trozos(progs, max(1, a.por_lamina))
            for j, pg in enumerate(gr):
                rg = figdir / f"Grid_{sem}_{nombre_seguro(esc)}_{j + 1}.png"
                fig_grid(pg, curs, meta.meta, per, actual, (lo, hi), tk, col, rg)
                lam.append(("imagen", f"Semestre {sem[0]}: programas de {esc}" + (f" ({j + 1} de {len(gr)})" if len(gr) > 1 else ""),
                            f"Cohortes {', '.join(per)}. Azul: {actual}{' (en curso)' if encurso else ''}. Rojo: meta{' de referencia' if ref else ''}. Título de cada panel: nuevos/meta de {actual}. Curvas semanales suavizadas.", rg))
        lam_sem.append(lam)

    anios = sorted({int(p[:4]) for p in brechas if p.endswith("10")} & {int(p[:4]) for p in brechas if p.endswith("20")})
    lam_brecha = []
    if anios:
        y = anios[-1]; rc = figdir / f"Comparativo_brecha_{y}.png"; fig_comparativo(y, brechas[f"{y}10"], brechas[f"{y}20"], rc)
        lam_brecha.append(("imagen", f"Brecha frente a la meta {y}: semestre 1 y semestre 2", "Nuevos de primera vez por encima (+) o por debajo (-) de la meta de cada programa.", rc))
    else:
        for c, g in brechas.items():
            rc = figdir / f"Brecha_{c}.png"; fig_brecha(c, g, corte[c], rc)
            lam_brecha.append(("imagen", f"Brecha frente a la meta, cohorte {c}", resumen_brecha(g), rc))

    # permanencia: solo entre periodos cerrados
    keys = [p for p in files if p not in en_curso]; ids = {p: set(DEP[p].Id) for p in keys}; filas, rows = [], []
    for i, c in enumerate(keys):
        base = set(NV[c].Id)
        if not base or i + 1 >= len(keys):
            continue
        filas.append((c, len(base), [len(base & ids[keys[i + j]]) / len(base) for j in range(1, len(keys) - i)]))
        nx = ids[keys[i + 1]]
        for prog, gg in NV[c].groupby("Programa"):
            sp = set(gg.Id); rows.append({"Cohorte": c, "Programa": prog, "n": len(sp), "Persisten_+1": len(sp & nx), "%_+1": round(100 * len(sp & nx) / len(sp), 1)})
    lam_perm = []
    if filas:
        rp = figdir / "Persistencia_cohortes.png"; fig_persistencia(filas, rp, titulos=False)
        hojas["Persistencia_total"] = pd.DataFrame([{"Cohorte": f[0], "n": f[1], **{f"+{j + 1}": round(100 * v, 1) for j, v in enumerate(f[2])}} for f in filas]).set_index("Cohorte")
        hojas["Persistencia_programa"] = pd.DataFrame(rows).set_index("Cohorte")
        lineas.append("Permanencia a +1 semestre: " + ", ".join(f"{f[0]} {100 * f[2][0]:.0f}%" for f in filas))
        p1 = [100 * f[2][0] for f in filas]
        kpis.append({"label": "Permanencia a +1 semestre", "valor": f"{p1[-1]:.0f}%", "sub": f"cohorte {filas[-1][0]}. Cohortes anteriores: {min(p1[:-1]):.0f}% a {max(p1[:-1]):.0f}%." if len(p1) > 1 else f"cohorte {filas[-1][0]}."})
        lam_perm.append(("imagen", "Permanencia de los nuevos por cohorte", "Porcentaje que vuelve a aparecer con pago en los semestres siguientes. No distingue deserción de grado, pausa o cambio de programa.", rp))

    perf = []; allc = pd.concat([NV[p].Ciudad.fillna("SIN DATO") for p in files]).value_counts(); top = list(allc.index[:6])
    for p in files:
        nv = NV[p]
        if nv.empty:
            continue
        ciu = nv.Ciudad.fillna("SIN DATO").where(lambda s: s.isin(top), "OTRAS")
        edad = (ancla(p) - nv.Fecha_nacimento).dt.days / 365.25
        ed = pd.cut(edad, [0, 17, 18, 19, 21, 200], right=False, labels=["<17", "17", "18", "19-20", "21+"])
        for var, s in (("Ciudad", ciu), ("Rango", nv.Rango.fillna("SIN DATO")), ("Edad (aprox.)", ed.astype(str))):
            for cat, k in s.value_counts().items():
                perf.append({"Periodo": p, "Variable": var, "Categoria": cat, "n": int(k), "%": round(100 * k / len(nv), 1)})
    hojas["Perfil_ingreso"] = pd.DataFrame(perf).set_index("Periodo")

    guardar_excel(hojas, out / "Matriculas_resumen.xlsx")
    (out / "resumen.txt").write_text("\n".join(lineas) + "\n", encoding="utf-8")

    # ---------------- documento
    cq = hojas["Calidad_datos"].copy(); cq["En_curso"] = cq["En_curso"].map({True: "Sí", False: "No"})
    cq = cq.drop(columns=["Filas"]).rename(columns={"En_curso": "En curso", "PV_bruto": "Primera vez (bruto)", "PV_devoluciones": "Devoluciones",
                                                    "PV_Id_duplicados": "Id duplicados", "PV_usado": "Usados", "PV_fuera_de_metas": "Fuera de metas", "Ultimo_pago": "Último pago"})
    cq["Último pago"] = cq["Último pago"].astype(str)
    notas += ["Definición: primera vez, sin devoluciones ni duplicados (una fila por Id y programa). Con --bruto se reproduce el conteo anterior.",
              "Comparar a igual día mide ritmo y no proyecta el cierre: en las cohortes 2024-1 a 2026-1 las diferencias de noviembre y diciembre se redujeron en enero."]
    ult = max(corte.values())
    lam = [("portada", "Matrículas de primera vez", [f"Monitoreo por cohortes, {min(files)} a {max(files)}",
                                                  (f"Fecha de reporte: {fecha_es(rep_fecha)}", {"bold": True, "color": INK}), f"Datos al {fecha_es(ult)}"]),
           ("resumen", "Resumen", f"Reportes BAN (Retención Levantada) al {fecha_es(ult)}", kpis[:4], notas)]
    if seg_lamina:
        lam.append(seg_lamina)
    lam += lam_proy
    if bts:
        metodo = [("Método: fracción acumulada histórica", {"size": 18, "bold": True}),
                  "1.  Para cada cohorte cerrada del mismo semestre se calcula qué fracción de su cierre ya estaba pagada ese mismo día del ciclo.",
                  "2.  Se divide lo que lleva la cohorte actual entre esa fracción. Cada cohorte de referencia produce una estimación del cierre.",
                  "3.  El rango es el mínimo y el máximo de las estimaciones. El valor central es la mediana.",
                  f"4.  Prueba retrospectiva: cada cohorte cerrada se proyecta con las demás y se compara con su cierre real. La proyección se declara confiable desde el día en que el error máximo deja de superar {100 * UMBRAL:.0f}%.",
                  ("Límites", {"size": 18, "bold": True}),
                  "Con tres cohortes el rango es indicativo y no un intervalo de confianza.",
                  "Supone que el calendario de pagos se repite. Si cambian las fechas de matrícula o las políticas de pago, la proyección pierde validez.",
                  "Se calcula solo para el total de la UTB, porque con cifras por programa la prueba retrospectiva es demasiado ruidosa.",
                  "La curva diaria de cada cohorte y el error de la prueba por día están en el Excel (hojas Curva_UTB_diaria y Proyeccion_prueba)."]
        lam.append(("texto", "Cómo se calcula la proyección de cierre", "Se usa solo con cohortes cerradas del mismo semestre y se valida contra ellas.", metodo))
        rbt = figdir / "Proyeccion_prueba.png"; fig_backtest(bts, rbt)
        par = [("Qué tan bien funcionó", {"size": 14, "bold": True})]
        for b in bts:
            f0 = ancla(b["ref"]); fila = []
            for dd in (30, 60, 90):
                if dd in b["mx"].index:
                    fila.append(f"+{dd} días ({fecha_es(f0 + pd.Timedelta(days=dd))}): {fmt_err(b['mx'].loc[dd])}")
            par.append((f"Semestre {b['sem'][0]}, error máximo", {"size": 12, "bold": True, "color": SUB}))
            par += [(x, {"size": 12}) for x in fila]
            par.append((f"Confiable desde {fecha_es(f0 + pd.Timedelta(days=b['dstar']))}." if b["dstar"] is not None else "No alcanzó el umbral de forma sostenida.", {"size": 12, "color": ACC}))
        lam.append(("mixto", "Prueba retrospectiva de la proyección", "Qué habría proyectado el método en cada fecha para las cohortes que ya cerraron.", rbt, par, "izq", 0.62))
    for l in lam_sem:
        lam += l
    lam += lam_brecha + lam_perm
    lam.append(("tabla", "Calidad de los datos por periodo", "Lo que se descontó para llegar al conteo de primera vez.", cq.rename_axis("Periodo"), 11.5))
    pie = f"Fuente: reportes BAN, Retención Levantada. Primera vez, " + ("conteo bruto" if a.bruto else "sin devoluciones ni duplicados") + f". Datos al {fecha_es(ult)}. Reporte del {fecha_es(rep_fecha)}."
    ruta_deck = out / f"Matriculas_{rep_fecha:%Y%m%d}.pptx"
    if generar_deck(ruta_deck, lam, pie):
        print("Documento:", ruta_deck)
    if not a.figuras:
        shutil.rmtree(figdir, ignore_errors=True)
    print("\n".join(lineas)); print("Listo:", out)

if __name__ == "__main__":
    main()
