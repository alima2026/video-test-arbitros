import streamlit as st
import pandas as pd
import json
import re
import smtplib
from email.message import EmailMessage
from pathlib import Path
from datetime import datetime
from io import BytesIO
import zipfile
import shutil

st.set_page_config(page_title="Video Test Árbitros", page_icon="⚽", layout="wide")

PREGUNTAS_FILE = Path("preguntas.json")
RESULTADOS_DIR = Path("resultados")
RESULTADOS_DIR.mkdir(exist_ok=True)

RESUMEN_CSV = RESULTADOS_DIR / "resultados_resumen.csv"
DETALLE_CSV = RESULTADOS_DIR / "resultados_detalle.csv"
EXCEL_FILE = RESULTADOS_DIR / "resultados_videotest.xlsx"

DECISIONES = ["No falta", "Tiro libre directo", "Tiro libre indirecto", "Penal"]
SANCIONES = ["No tarjeta", "Amonestación", "Expulsión"]


def get_secret(nombre, default=None):
    try:
        return st.secrets.get(nombre, default)
    except Exception:
        return default


def clave_admin():
    return get_secret("ADMIN_PASSWORD", "admin123")


def email_valido(email):
    return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", str(email).strip()) is not None


def email_configurado():
    return bool(
        get_secret("SMTP_SERVER", "") and
        get_secret("SMTP_USER", "") and
        get_secret("SMTP_PASSWORD", "") and
        get_secret("EMAIL_FROM", "")
    )


def cargar():
    with open(PREGUNTAS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar(data):
    with open(PREGUNTAS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def nivel(porc):
    if porc >= 90:
        return "Excelente"
    if porc >= 80:
        return "Muy bueno"
    if porc >= 70:
        return "Bueno"
    if porc >= 60:
        return "Regular"
    return "Debe reforzar"


def faltan_respuestas(data):
    faltan = []
    for v in data["videos"]:
        if v.get("decision_correcta", "") not in DECISIONES:
            faltan.append(f'{v.get("titulo","")}: falta decisión técnica')
        if v.get("sancion_correcta", "") not in SANCIONES:
            faltan.append(f'{v.get("titulo","")}: falta sanción disciplinaria')
    return faltan


def leer_csv(path):
    return pd.read_csv(path) if path.exists() else pd.DataFrame()


def normalizar_archivo(nombre):
    return str(nombre).lower().replace(" ", "").replace("-", "_")


def buscar_video(ruta_configurada, video_id):
    ruta = Path(str(ruta_configurada))
    if ruta.exists():
        return ruta

    carpeta = ruta.parent if ruta.parent != Path(".") else Path("videos")
    if not carpeta.exists():
        return None

    objetivo = normalizar_archivo(ruta.name)
    for archivo in carpeta.glob("*.mp4"):
        if normalizar_archivo(archivo.name) == objetivo:
            return archivo

    numero = str(video_id).lower().replace("v", "").strip()
    posibles = {
        normalizar_archivo(f"video_{numero}.mp4"),
        normalizar_archivo(f"Video_{numero}.mp4"),
        normalizar_archivo(f"Video_ {numero}.mp4"),
    }
    for archivo in carpeta.glob("*.mp4"):
        if normalizar_archivo(archivo.name) in posibles:
            return archivo
    return None


def guardar_seguro_csv(df, path, base_name):
    try:
        df.to_csv(path, index=False, encoding="utf-8-sig")
    except PermissionError:
        nuevo = RESULTADOS_DIR / f"{base_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df.to_csv(nuevo, index=False, encoding="utf-8-sig")
        st.warning(f"{path.name} estaba abierto. Guardé copia: {nuevo.name}")


def guardar_seguro_excel(resumen, detalle):
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            resumen.to_excel(writer, index=False, sheet_name="Resumen")
            detalle.to_excel(writer, index=False, sheet_name="Detalle")
    except PermissionError:
        nuevo = RESULTADOS_DIR / f"resultados_videotest_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(nuevo, engine="openpyxl") as writer:
            resumen.to_excel(writer, index=False, sheet_name="Resumen")
            detalle.to_excel(writer, index=False, sheet_name="Detalle")
        st.warning(f"El Excel estaba abierto. Guardé copia: {nuevo.name}")


def corregir(data, respuestas):
    detalle = []
    puntos = 0
    total = 0

    for v in data["videos"]:
        pares = [
            ("Decisión técnica", respuestas[f'{v["id"]}_decision'], v["decision_correcta"], 2),
            ("Sanción disciplinaria", respuestas[f'{v["id"]}_sancion'], v["sancion_correcta"], 2),
        ]

        for pregunta, resp, correcta, pts in pares:
            ok = resp == correcta
            total += pts
            puntos += pts if ok else 0

            detalle.append({
                "tema": v.get("tema", ""),
                "subtema": v.get("subtema", ""),
                "video": v.get("titulo", ""),
                "pregunta": pregunta,
                "respuesta_usuario": resp,
                "respuesta_correcta": correcta,
                "correcta": "Sí" if ok else "No",
                "puntos": pts,
                "obtenido": pts if ok else 0,
                "criterio_admin": v.get("criterio_admin", ""),
                "explicacion_admin": v.get("explicacion_admin", ""),
            })

    porc = round((puntos / total) * 100, 2) if total else 0
    return puntos, total, porc, nivel(porc), detalle


def resumen_temas(detalle):
    df = pd.DataFrame(detalle)
    if df.empty:
        return pd.DataFrame()
    res = df.groupby("tema", as_index=False).agg(
        puntos_obtenidos=("obtenido", "sum"),
        puntos_totales=("puntos", "sum")
    )
    res["porcentaje"] = (res["puntos_obtenidos"] / res["puntos_totales"] * 100).round(2)
    res["nivel"] = res["porcentaje"].apply(nivel)
    return res


def videos_a_reforzar(detalle):
    """
    Devuelve una tabla formativa para el participante.
    No muestra las respuestas correctas, solo qué videos/aspectos debe revisar.
    """
    df = pd.DataFrame(detalle)

    if df.empty:
        return pd.DataFrame(columns=[
            "video", "tema", "subtema", "aspecto_a_revisar",
            "rendimiento_video", "prioridad"
        ])

    df["ok"] = df["correcta"].astype(str).str.lower().isin(["sí", "si", "true", "1"])

    filas = []
    for video, g in df.groupby("video", sort=False):
        errores = g[~g["ok"]].copy()

        if errores.empty:
            continue

        aspectos = []
        for pregunta in errores["pregunta"].tolist():
            if "Decisión" in str(pregunta):
                aspectos.append("Revisar decisión técnica")
            elif "Sanción" in str(pregunta):
                aspectos.append("Revisar sanción disciplinaria")
            else:
                aspectos.append(f"Revisar {pregunta}")

        obtenidos = int(g["obtenido"].sum())
        total = int(g["puntos"].sum())
        porc_video = round((obtenidos / total) * 100, 2) if total else 0

        prioridad = "Alta" if len(errores) >= 2 else "Media"

        filas.append({
            "video": video,
            "tema": g["tema"].iloc[0] if "tema" in g.columns else "",
            "subtema": g["subtema"].iloc[0] if "subtema" in g.columns else "",
            "aspecto_a_revisar": " / ".join(sorted(set(aspectos))),
            "rendimiento_video": f"{obtenidos}/{total} ({porc_video}%)",
            "prioridad": prioridad,
        })

    return pd.DataFrame(filas)




def analisis_individual_arbitro(detalle):
    df = pd.DataFrame(detalle)
    if df.empty:
        return {
            "fortalezas": [],
            "a_reforzar": [],
            "videos_criticos": [],
            "errores_decision": 0,
            "errores_sancion": 0,
            "total_decision": 0,
            "total_sancion": 0,
            "acierto_decision": 0,
            "acierto_sancion": 0,
        }

    df["ok"] = df["correcta"].astype(str).str.lower().isin(["sí", "si", "true", "1"])

    decision = df[df["pregunta"].astype(str).str.contains("Decisión", case=False, na=False)]
    sancion = df[df["pregunta"].astype(str).str.contains("Sanción", case=False, na=False)]

    errores_decision = int((~decision["ok"]).sum()) if not decision.empty else 0
    errores_sancion = int((~sancion["ok"]).sum()) if not sancion.empty else 0
    total_decision = int(len(decision))
    total_sancion = int(len(sancion))

    acierto_decision = round((decision["ok"].sum() / len(decision)) * 100, 2) if len(decision) else 0
    acierto_sancion = round((sancion["ok"].sum() / len(sancion)) * 100, 2) if len(sancion) else 0

    por_tema = df.groupby("tema", as_index=False).agg(
        respuestas=("ok", "count"),
        aciertos=("ok", "sum"),
        errores=("ok", lambda s: (~s).sum())
    )
    por_tema["porcentaje_acierto"] = (por_tema["aciertos"] / por_tema["respuestas"] * 100).round(2)

    fortalezas = por_tema[por_tema["porcentaje_acierto"] >= 80].sort_values("porcentaje_acierto", ascending=False)
    a_reforzar = por_tema[por_tema["porcentaje_acierto"] < 80].sort_values("porcentaje_acierto", ascending=True)

    por_video = df.groupby(["video", "tema", "subtema"], as_index=False).agg(
        respuestas=("ok", "count"),
        aciertos=("ok", "sum"),
        errores=("ok", lambda s: (~s).sum()),
        puntos_obtenidos=("obtenido", "sum"),
        puntos_totales=("puntos", "sum"),
    )
    por_video["porcentaje_acierto"] = (por_video["aciertos"] / por_video["respuestas"] * 100).round(2)

    videos_criticos = por_video[por_video["errores"] > 0].sort_values(
        ["errores", "porcentaje_acierto"], ascending=[False, True]
    )

    return {
        "fortalezas": fortalezas["tema"].tolist(),
        "a_reforzar": a_reforzar["tema"].tolist(),
        "videos_criticos": videos_criticos.to_dict(orient="records"),
        "errores_decision": errores_decision,
        "errores_sancion": errores_sancion,
        "total_decision": total_decision,
        "total_sancion": total_sancion,
        "acierto_decision": acierto_decision,
        "acierto_sancion": acierto_sancion,
    }


def conclusion_general_individual(participante, puntos, total, porc, niv, detalle):
    nombre = participante.get("nombre", "El participante")
    analisis = analisis_individual_arbitro(detalle)

    if porc >= 90:
        encabezado = f"{nombre} alcanzó un rendimiento excelente. Demuestra un criterio sólido en la lectura técnica y disciplinaria de las jugadas observadas."
    elif porc >= 80:
        encabezado = f"{nombre} alcanzó un rendimiento muy bueno. En líneas generales interpreta correctamente las situaciones, aunque presenta algunos puntos puntuales para revisar."
    elif porc >= 70:
        encabezado = f"{nombre} alcanzó un rendimiento bueno. El resultado es positivo, pero se recomienda profundizar en las jugadas donde existieron dudas para consolidar el criterio."
    elif porc >= 60:
        encabezado = f"{nombre} alcanzó un rendimiento regular. Se observan aspectos importantes para reforzar, especialmente en la identificación del tipo de infracción y/o la sanción disciplinaria."
    else:
        encabezado = f"{nombre} debe reforzar el análisis de las jugadas. Es recomendable volver a revisar los clips marcados y trabajar los criterios técnicos antes de una nueva evaluación."

    lineas = [
        encabezado,
        "",
        f"Resultado general: {niv} ({porc}%). Puntaje obtenido: {puntos}/{total}.",
        "",
        "Lectura técnica y disciplinaria:",
        f"- Decisión técnica: {analisis['acierto_decision']}% de acierto ({analisis['errores_decision']} errores sobre {analisis['total_decision']} jugadas).",
        f"- Sanción disciplinaria: {analisis['acierto_sancion']}% de acierto ({analisis['errores_sancion']} errores sobre {analisis['total_sancion']} jugadas).",
    ]

    if analisis["fortalezas"]:
        lineas.append("")
        lineas.append("Fortalezas detectadas:")
        for tema in analisis["fortalezas"][:4]:
            lineas.append(f"- Buen rendimiento en {tema}.")

    if analisis["a_reforzar"]:
        lineas.append("")
        lineas.append("Aspectos generales a reforzar:")
        for tema in analisis["a_reforzar"][:4]:
            lineas.append(f"- Reforzar criterio en {tema}.")

    if analisis["videos_criticos"]:
        lineas.append("")
        lineas.append("Videos sugeridos para volver a analizar:")
        for r in analisis["videos_criticos"][:6]:
            errores = int(r.get("errores", 0))
            prioridad = "alta" if errores >= 2 else "media"
            lineas.append(
                f"- {r.get('video','')} ({r.get('tema','')}): revisar nuevamente. "
                f"Prioridad {prioridad}. Rendimiento del video: "
                f"{int(r.get('puntos_obtenidos', 0))}/{int(r.get('puntos_totales', 0))}."
            )
    else:
        lineas.append("")
        lineas.append("No se detectaron videos con errores. Se recomienda mantener el criterio aplicado.")

    lineas.append("")
    lineas.append("Esta conclusión es formativa: orienta qué revisar, sin mostrar las respuestas correctas del banco.")
    return "\n".join(lineas)


def cuerpo_mail_resultado(participante, puntos, total, porc, niv, detalle):
    nombre = participante.get("nombre", "Participante")
    res_tema = resumen_temas(detalle)
    reforzar = videos_a_reforzar(detalle)

    lineas = [
        f"Estimado/a {nombre}:",
        "",
        "Se informa el resultado de su Video Test de Árbitros.",
        "",
        f"Resultado: {niv}",
        f"Puntaje obtenido: {puntos} / {total}",
        f"Porcentaje: {porc}%",
        "",
        "Conclusión general:",
        conclusion_general_individual(participante, puntos, total, porc, niv, detalle),
        "",
    ]

    if not res_tema.empty:
        lineas.append("Resumen por tema:")
        for _, r in res_tema.iterrows():
            lineas.append(f"- {r['tema']}: {r['puntos_obtenidos']} / {r['puntos_totales']} ({r['porcentaje']}%) - {r['nivel']}")
        lineas.append("")

    if reforzar.empty:
        lineas.append("Videos a reforzar / analizar:")
        lineas.append("- No se detectaron videos con errores. Mantener el criterio aplicado.")
        lineas.append("")
    else:
        lineas.append("Videos a reforzar / analizar:")
        for _, r in reforzar.iterrows():
            lineas.append(
                f"- {r['video']} | {r['tema']} | {r['aspecto_a_revisar']} | "
                f"Rendimiento: {r['rendimiento_video']} | Prioridad: {r['prioridad']}"
            )
        lineas.append("")

    lineas.extend([
        "Este resultado tiene finalidad formativa.",
        "Las respuestas correctas quedan disponibles únicamente para el administrador.",
        "",
        "Departamento de Arbitraje",
        "Comisión Juvenil AUF",
    ])
    return "\n".join(lineas)


def enviar_mail(destinatario, asunto, cuerpo):
    smtp_server = get_secret("SMTP_SERVER", "")
    smtp_port = int(get_secret("SMTP_PORT", 587))
    smtp_user = get_secret("SMTP_USER", "")
    smtp_password = get_secret("SMTP_PASSWORD", "")
    email_from = get_secret("EMAIL_FROM", smtp_user)

    if not smtp_server or not smtp_user or not smtp_password or not email_from:
        return False, "Falta configurar SMTP en Secrets."

    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = email_from
    msg["To"] = destinatario
    msg.set_content(cuerpo)

    try:
        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as smtp:
            smtp.starttls()
            smtp.login(smtp_user, smtp_password)
            smtp.send_message(msg)
        return True, "Enviado"
    except Exception as e:
        return False, str(e)


def enviar_resultado_participante(participante, puntos, total, porc, niv, detalle):
    email = participante.get("email", "").strip()
    if not email_valido(email):
        return False, "Correo inválido"
    cuerpo = cuerpo_mail_resultado(participante, puntos, total, porc, niv, detalle)
    return enviar_mail(email, "Resultado Video Test de Árbitros", cuerpo)


def guardar_resultado(participante, puntos, total, porc, niv, detalle, estado_mail):
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    fila = {
        "fecha": fecha,
        **participante,
        "puntaje_obtenido": puntos,
        "puntaje_total": total,
        "porcentaje": porc,
        "nivel": niv,
        "mail_resultado": estado_mail,
    }

    resumen_nuevo = pd.DataFrame([fila])
    detalle_nuevo = pd.DataFrame(detalle)

    for k, v in fila.items():
        detalle_nuevo.insert(0, k, v)

    resumen = pd.concat([leer_csv(RESUMEN_CSV), resumen_nuevo], ignore_index=True)
    detalle_final = pd.concat([leer_csv(DETALLE_CSV), detalle_nuevo], ignore_index=True)

    guardar_seguro_csv(resumen, RESUMEN_CSV, "resultados_resumen")
    guardar_seguro_csv(detalle_final, DETALLE_CSV, "resultados_detalle")
    guardar_seguro_excel(resumen, detalle_final)


def excel_en_memoria(resumen, detalle):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        resumen.to_excel(writer, index=False, sheet_name="Resumen")
        detalle.to_excel(writer, index=False, sheet_name="Detalle")
    out.seek(0)
    return out.getvalue()


def zip_resultados():
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for f in [RESUMEN_CSV, DETALLE_CSV, EXCEL_FILE]:
            if f.exists():
                z.write(f, f.name)
    out.seek(0)
    return out.getvalue()


def borrar_resultados_con_backup():
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = RESULTADOS_DIR / f"backup_{stamp}"
    backup.mkdir(exist_ok=True)

    copiados = []
    borrados = []
    for f in [RESUMEN_CSV, DETALLE_CSV, EXCEL_FILE]:
        if f.exists():
            shutil.copy2(f, backup / f.name)
            copiados.append(f.name)
            f.unlink()
            borrados.append(f.name)
    return backup, copiados, borrados


def analisis_preguntas(detalle):
    if detalle.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df = detalle.copy()
    df["ok"] = df["correcta"].astype(str).str.lower().isin(["sí", "si", "true", "1"])
    df["pregunta_id"] = df["video"].astype(str) + " - " + df["pregunta"].astype(str)

    por_pregunta = df.groupby(["pregunta_id", "tema", "subtema", "video", "pregunta"], as_index=False).agg(
        participantes=("email", "count"),
        aciertos=("ok", "sum"),
        errores=("ok", lambda s: (~s).sum())
    )
    por_pregunta["porcentaje_acierto"] = (por_pregunta["aciertos"] / por_pregunta["participantes"] * 100).round(2)
    por_pregunta["porcentaje_error"] = (por_pregunta["errores"] / por_pregunta["participantes"] * 100).round(2)
    por_pregunta["division"] = (100 - (por_pregunta["porcentaje_acierto"] - 50).abs() * 2).round(2)

    por_tema = df.groupby(["tema", "subtema"], as_index=False).agg(
        respuestas=("ok", "count"),
        aciertos=("ok", "sum"),
        errores=("ok", lambda s: (~s).sum())
    )
    por_tema["porcentaje_acierto"] = (por_tema["aciertos"] / por_tema["respuestas"] * 100).round(2)
    por_tema["porcentaje_error"] = (por_tema["errores"] / por_tema["respuestas"] * 100).round(2)

    return (
        por_pregunta.sort_values(["porcentaje_error", "errores"], ascending=[False, False]),
        por_pregunta.sort_values(["porcentaje_acierto", "aciertos"], ascending=[False, False]),
        por_pregunta.sort_values("division", ascending=False),
        por_tema.sort_values("porcentaje_error", ascending=False),
    )


def diagnostico_grupal(resumen, detalle):
    if resumen.empty or detalle.empty:
        return "Todavía no hay datos suficientes."

    mas_falladas, mas_aciertos, mas_divididas, por_tema = analisis_preguntas(detalle)
    participantes = resumen["email"].nunique() if "email" in resumen.columns else len(resumen)
    promedio = resumen["porcentaje"].mean()

    lineas = [
        f"Participantes evaluados: {participantes}.",
        f"Promedio general: {promedio:.2f}%.",
    ]

    if not mas_falladas.empty:
        r = mas_falladas.iloc[0]
        lineas.append(f"Pregunta más fallada: {r['pregunta_id']} con {r['porcentaje_error']}% de error.")

    if not mas_aciertos.empty:
        r = mas_aciertos.iloc[0]
        lineas.append(f"Pregunta con más aciertos: {r['pregunta_id']} con {r['porcentaje_acierto']}% de acierto.")

    if not mas_divididas.empty:
        r = mas_divididas.iloc[0]
        lineas.append(f"Pregunta más dividida: {r['pregunta_id']} ({r['porcentaje_acierto']}% acierto / {r['porcentaje_error']}% error).")

    if not por_tema.empty:
        r = por_tema.iloc[0]
        lineas.append(f"Tema a reforzar: {r['tema']} - {r['subtema']} ({r['porcentaje_error']}% de error).")

    return "\n".join(lineas)


def excel_analisis(resumen, detalle):
    mas_falladas, mas_aciertos, mas_divididas, por_tema = analisis_preguntas(detalle)
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        resumen.to_excel(writer, index=False, sheet_name="Resumen")
        detalle.to_excel(writer, index=False, sheet_name="Detalle")
        mas_falladas.to_excel(writer, index=False, sheet_name="Mas falladas")
        mas_aciertos.to_excel(writer, index=False, sheet_name="Mas aciertos")
        mas_divididas.to_excel(writer, index=False, sheet_name="Mas divididas")
        por_tema.to_excel(writer, index=False, sheet_name="Por tema")
    out.seek(0)
    return out.getvalue()


def reenviar_resultado(fila, detalle):
    email = str(fila.get("email", "")).strip()
    if not email_valido(email):
        return False, "Correo inválido"

    participante = {
        "nombre": fila.get("nombre", ""),
        "email": email,
        "categoria": fila.get("categoria", ""),
        "institucion": fila.get("institucion", ""),
    }

    filtro = detalle[
        (detalle["email"].astype(str) == str(email)) &
        (detalle["fecha"].astype(str) == str(fila.get("fecha", "")))
    ].copy()

    det = filtro.to_dict(orient="records") if not filtro.empty else []
    cuerpo = cuerpo_mail_resultado(
        participante,
        fila.get("puntaje_obtenido", 0),
        fila.get("puntaje_total", 0),
        fila.get("porcentaje", 0),
        fila.get("nivel", ""),
        det
    )

    return enviar_mail(email, "Resultado Video Test de Árbitros", cuerpo)


def admin():
    st.header("Panel administrador")
    pwd = st.text_input("Clave administrador", type="password")

    if pwd != clave_admin():
        st.info("Clave local inicial: admin123")
        return

    data = cargar()
    resumen = leer_csv(RESUMEN_CSV)
    detalle = leer_csv(DETALLE_CSV)

    st.success("Acceso administrador correcto")

    tab_config, tab_resultados, tab_analisis, tab_correos, tab_videos = st.tabs([
        "Configuración",
        "Resultados / guardar / borrar",
        "Análisis grupal",
        "Correos",
        "Chequeo de videos",
    ])

    with tab_config:
        st.subheader("Configuración de respuestas correctas")

        df = pd.DataFrame(data["videos"])
        df_edit = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            disabled=["id"],
            column_config={
                "decision_correcta": st.column_config.SelectboxColumn("decision_correcta", options=[""] + DECISIONES),
                "sancion_correcta": st.column_config.SelectboxColumn("sancion_correcta", options=[""] + SANCIONES),
            },
        )

        if st.button("Guardar configuración"):
            data["videos"] = df_edit.to_dict(orient="records")
            guardar(data)
            st.success("Configuración guardada")
            st.rerun()

        faltan = faltan_respuestas(data)
        if faltan:
            st.warning("Faltan respuestas correctas:")
            for x in faltan:
                st.write(f"- {x}")
        else:
            st.success("Banco completo.")

    with tab_resultados:
        st.subheader("Resultados generales")

        if resumen.empty:
            st.warning("Todavía no hay resultados.")
        else:
            st.dataframe(resumen, use_container_width=True)
            st.subheader("Detalle de respuestas")
            st.dataframe(detalle, use_container_width=True)

            c1, c2, c3, c4 = st.columns(4)
            c1.download_button("Descargar resumen CSV", RESUMEN_CSV.read_bytes(), "resultados_resumen.csv", mime="text/csv")
            c2.download_button("Descargar detalle CSV", DETALLE_CSV.read_bytes(), "resultados_detalle.csv", mime="text/csv")
            c3.download_button("Descargar Excel", excel_en_memoria(resumen, detalle), "resultados_videotest.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            c4.download_button("Guardar backup ZIP", zip_resultados(), f"backup_resultados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip", mime="application/zip")

        st.divider()
        st.subheader("Borrar resultados para nuevo video test")
        st.warning("Antes de borrar, descargá el Excel o el backup ZIP. La app también creará un backup interno antes de borrar.")
        conf = st.text_input("Para borrar escribí exactamente: BORRAR RESULTADOS")

        if st.button("Borrar resultados", type="primary"):
            if conf.strip() == "BORRAR RESULTADOS":
                backup, copiados, borrados = borrar_resultados_con_backup()
                st.success(f"Resultados borrados. Backup interno: {backup}")
                st.write("Archivos borrados:", borrados)
                st.rerun()
            else:
                st.error("No se borró nada. Debés escribir exactamente: BORRAR RESULTADOS")

    with tab_analisis:
        st.subheader("Análisis grupal")

        if resumen.empty or detalle.empty:
            st.warning("Todavía no hay resultados suficientes.")
        else:
            participantes = resumen["email"].nunique() if "email" in resumen.columns else len(resumen)
            promedio = resumen["porcentaje"].mean()
            aprobados = int((resumen["porcentaje"] >= 70).sum())
            reforzar = int((resumen["porcentaje"] < 60).sum())

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Participantes", participantes)
            c2.metric("Promedio", f"{promedio:.2f}%")
            c3.metric("70% o más", aprobados)
            c4.metric("Debe reforzar", reforzar)

            mas_falladas, mas_aciertos, mas_divididas, por_tema = analisis_preguntas(detalle)

            st.subheader("Preguntas más falladas")
            st.dataframe(mas_falladas, use_container_width=True)

            st.subheader("Preguntas con más aciertos")
            st.dataframe(mas_aciertos, use_container_width=True)

            st.subheader("Preguntas más divididas")
            st.caption("Dividida = cuando el grupo queda cerca de 50% acierto y 50% error.")
            st.dataframe(mas_divididas, use_container_width=True)

            st.subheader("Rendimiento por tema / subtema")
            st.dataframe(por_tema, use_container_width=True)

            st.subheader("Diagnóstico automático grupal")
            st.text_area("Diagnóstico grupal", diagnostico_grupal(resumen, detalle), height=180)

            st.subheader("Conclusión individual por árbitro")
            opciones_ind = []
            for idx, r in resumen.iterrows():
                opciones_ind.append(f"{idx} - {r.get('nombre','')} - {r.get('email','')} - {r.get('fecha','')}")

            seleccionado_ind = st.selectbox("Seleccionar árbitro / resultado", opciones_ind, key="sel_conclusion_individual")
            idx_ind = int(seleccionado_ind.split(" - ")[0])
            fila_ind = resumen.iloc[idx_ind]

            det_ind = detalle[
                (detalle["email"].astype(str) == str(fila_ind.get("email", ""))) &
                (detalle["fecha"].astype(str) == str(fila_ind.get("fecha", "")))
            ].copy()

            participante_ind = {
                "nombre": fila_ind.get("nombre", ""),
                "email": fila_ind.get("email", ""),
                "categoria": fila_ind.get("categoria", ""),
                "institucion": fila_ind.get("institucion", ""),
            }

            conclusion_ind = conclusion_general_individual(
                participante_ind,
                fila_ind.get("puntaje_obtenido", 0),
                fila_ind.get("puntaje_total", 0),
                fila_ind.get("porcentaje", 0),
                fila_ind.get("nivel", ""),
                det_ind.to_dict(orient="records")
            )

            st.text_area("Conclusión individual", conclusion_ind, height=280)

            st.download_button(
                "Descargar análisis grupal Excel",
                excel_analisis(resumen, detalle),
                f"analisis_grupal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with tab_correos:
        st.subheader("Envío de resultados por mail")
        st.write("Remitente configurado:")
        st.code(get_secret("EMAIL_FROM", "No configurado"))

        if email_configurado():
            st.success("SMTP configurado.")
        else:
            st.error("SMTP no configurado. Agregá los datos en Secrets de Streamlit.")
            st.code(
                'ADMIN_PASSWORD = "admin123"\n'
                'EMAIL_FROM = "arbitraje.comjuvenil@auf.org.uy"\n'
                'SMTP_SERVER = "smtp.gmail.com"\n'
                'SMTP_PORT = 587\n'
                'SMTP_USER = "arbitraje.comjuvenil@auf.org.uy"\n'
                'SMTP_PASSWORD = "AQUI_LA_CONTRASEÑA_O_APP_PASSWORD"\n'
                'SEND_EMAIL_AUTO = true'
            )

        if resumen.empty:
            st.warning("No hay resultados para enviar.")
        else:
            st.dataframe(resumen, use_container_width=True)

            opciones = []
            for idx, r in resumen.iterrows():
                opciones.append(f"{idx} - {r.get('nombre','')} - {r.get('email','')} - {r.get('fecha','')}")

            seleccionado = st.selectbox("Seleccionar resultado", opciones)

            if st.button("Enviar resultado seleccionado"):
                idx = int(seleccionado.split(" - ")[0])
                ok, msg = reenviar_resultado(resumen.iloc[idx], detalle)
                if ok:
                    st.success(f"Correo enviado a {resumen.iloc[idx].get('email')}")
                else:
                    st.error(f"No se pudo enviar: {msg}")

            confirmar = st.checkbox("Confirmo que quiero enviar el resultado a todos los correos del resumen.")

            if st.button("Enviar a todos"):
                if not confirmar:
                    st.error("Primero confirmá el envío a todos.")
                else:
                    enviados = 0
                    fallidos = []
                    for _, r in resumen.iterrows():
                        ok, msg = reenviar_resultado(r, detalle)
                        if ok:
                            enviados += 1
                        else:
                            fallidos.append({"nombre": r.get("nombre", ""), "email": r.get("email", ""), "error": msg})
                    st.success(f"Correos enviados: {enviados}")
                    if fallidos:
                        st.warning("Algunos fallaron:")
                        st.dataframe(pd.DataFrame(fallidos), use_container_width=True)

    with tab_videos:
        st.subheader("Chequeo de videos")
        chequeo = []
        for v in data["videos"]:
            ruta_real = buscar_video(v.get("archivo", ""), v.get("id", ""))
            chequeo.append({
                "video": v.get("titulo", ""),
                "ruta_configurada": v.get("archivo", ""),
                "encontrado": "Sí" if ruta_real else "No",
                "ruta_real": str(ruta_real) if ruta_real else "",
            })
        st.dataframe(pd.DataFrame(chequeo), use_container_width=True)


def test():
    data = cargar()
    faltan = faltan_respuestas(data)

    if faltan:
        st.error("El test no está habilitado. El administrador debe cargar las respuestas correctas reales.")
        with st.expander("Ver faltantes"):
            for x in faltan:
                st.write(f"- {x}")
        return

    st.header("Ingreso del participante")

    with st.form("form_participante"):
        nombre = st.text_input("Nombre y apellido *")
        email = st.text_input("Correo electrónico *", placeholder="ejemplo: nombre@gmail.com")
        categoria = st.selectbox("Categoría / rol", ["Árbitro", "Árbitro asistente", "Cuarto árbitro", "Instructor", "Otro"])
        institucion = st.text_input("Institución / departamento")
        aceptar = st.checkbox("Confirmo que realizaré el test de forma individual.")
        ok = st.form_submit_button("Comenzar test")

    if ok:
        if not nombre.strip():
            st.error("Debe ingresar nombre y apellido.")
            st.stop()
        if not email_valido(email):
            st.error("Debe ingresar un correo válido.")
            st.stop()
        if not aceptar:
            st.error("Debe confirmar que realizará el test de forma individual.")
            st.stop()

        st.session_state["datos_participante"] = {
            "nombre": nombre.strip(),
            "email": email.strip(),
            "categoria": categoria,
            "institucion": institucion.strip(),
        }
        st.session_state["test_habilitado"] = True
        st.rerun()

    if not st.session_state.get("test_habilitado", False):
        st.info("Completá tus datos y presioná Comenzar test.")
        return

    respuestas = {}
    st.header("Video test")

    with st.form("form_test"):
        for i, v in enumerate(data["videos"], start=1):
            st.subheader(f"Video {i}: {v['titulo']}")

            ruta_real = buscar_video(v.get("archivo", ""), v.get("id", ""))
            if ruta_real:
                st.video(str(ruta_real))
            else:
                st.warning(f"No encontré el video: {v.get('archivo', '')}")

            respuestas[f'{v["id"]}_decision'] = st.radio("1. Decisión técnica", DECISIONES, key=f'{v["id"]}_d', index=None) or ""
            respuestas[f'{v["id"]}_sancion'] = st.radio("2. Sanción disciplinaria", SANCIONES, key=f'{v["id"]}_s', index=None) or ""
            st.divider()

        enviar = st.form_submit_button("Enviar respuestas y ver calificación")

    if enviar:
        if any(not r for r in respuestas.values()):
            st.error("Faltan preguntas por responder. Revisá el test antes de enviar.")
            st.stop()

        puntos, total, porc, niv, detalle = corregir(data, respuestas)
        participante = st.session_state["datos_participante"]

        auto_mail = str(get_secret("SEND_EMAIL_AUTO", "true")).lower() in ["true", "1", "yes", "si", "sí"]
        estado_mail = "No enviado"

        if auto_mail and email_configurado():
            ok_mail, msg_mail = enviar_resultado_participante(participante, puntos, total, porc, niv, detalle)
            estado_mail = "Enviado" if ok_mail else f"Error: {msg_mail}"
        elif auto_mail:
            estado_mail = "No enviado: SMTP no configurado"
        else:
            estado_mail = "No enviado: automático desactivado"

        guardar_resultado(participante, puntos, total, porc, niv, detalle, estado_mail)

        st.success("Test enviado correctamente")

        st.markdown(f"## Resultado: {niv}")

        c1, c2, c3 = st.columns(3)
        c1.metric("Puntaje", f"{puntos} / {total}")
        c2.metric("Porcentaje", f"{porc}%")
        c3.metric("Calificación", niv)

        st.subheader("Conclusión general")
        conclusion = conclusion_general_individual(participante, puntos, total, porc, niv, detalle)
        st.info(conclusion)

        st.subheader("Resumen por tema")
        st.dataframe(resumen_temas(detalle), use_container_width=True)

        st.subheader("Videos a reforzar / analizar")
        refuerzo = videos_a_reforzar(detalle)

        if refuerzo.empty:
            st.success("No se detectaron videos con errores. Mantené el criterio aplicado.")
        else:
            st.info("Estos son los videos donde conviene reforzar o volver a analizar el criterio. No se muestran las respuestas correctas al participante.")
            st.dataframe(refuerzo, use_container_width=True)

            for _, r in refuerzo.iterrows():
                prioridad = "🔴 Prioridad alta" if r["prioridad"] == "Alta" else "🟡 Prioridad media"
                st.write(
                    f"{prioridad} — **{r['video']}** | {r['tema']} | "
                    f"{r['aspecto_a_revisar']} | Rendimiento: {r['rendimiento_video']}"
                )

        if estado_mail == "Enviado":
            st.success("Tu resultado fue enviado por correo electrónico.")
        else:
            st.info("El resultado quedó guardado. El administrador podrá enviarlo por correo.")

        st.info("Las respuestas correctas quedan solo para el administrador.")


st.title("⚽ Video Test Árbitros")
st.caption("Versión con conclusión general individual, resultado formativo, administración, análisis grupal y envío de correos.")

menu = st.sidebar.radio("Menú", ["Realizar test", "Administrador"])

if menu == "Administrador":
    admin()
else:
    test()
