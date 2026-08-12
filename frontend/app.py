from __future__ import annotations

import io
import importlib
import re
import sqlite3
import shutil
import sys
import unicodedata
import zipfile
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from docx import Document


ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
	sys.path.append(str(ROOT_DIR))

from backend.database import (
	actualizar_unidad,
	actualizar_usuario,
	buscar_alumno_por_rut_anio_mencion,
	crear_alumno,
	crear_unidad,
	crear_usuario,
	eliminar_alumno,
	eliminar_evaluacion,
	eliminar_retroalimentacion_local,
	eliminar_usuario_por_username,
	eliminar_usuario,
	eliminar_unidad,
	formatear_rut_chileno,
	guardar_evaluacion,
	guardar_mapeo_pregunta,
	guardar_puntajes_pregunta,
	guardar_retroalimentacion_local,
	inicializar_si_no_existe,
	cargar_indicadores_desde_df,
	extraer_indicadores_desde_pdf,
	listar_evaluaciones_unidad,
	listar_alumnos,
	obtener_historial_retroalimentaciones,
	obtener_puntajes_pregunta,
	listar_indicadores_unidad,
	listar_mapeo_evaluacion,
	listar_resultados_alumno,
	listar_unidades,
	obtener_usuarios,
	obtener_alumnos,
	obtener_brechas_alumno,
	obtener_unidad_por_id,
	obtener_usuario_por_credenciales,
	obtener_usuario_por_username,
	reemplazar_indicadores_unidad,
	reemplazar_mapeo_evaluacion,
	reemplazar_resultados_alumno,
	validar_rut_chileno,
)
from backend.orquestador import generar_plan_trabajo_alumno, generar_plan_trabajo_cohorte
from backend.pipeline_rag import cargar_documentos_docx, construir_vector_store


DOCENTE_DEFAULT = "TCL Pablo Varas Briones"
ANIOS_CRIM = ["I CRIM", "II CRIM", "III CRIM", "IV CRIM"]
MENCIONES = [
	"Comunicaciones e Informatica",
	"Sistemas de Armas",
	"Abastecimiento y Finanzas",
	"Geografia",
]


st.set_page_config(
	page_title="ACAPOMIL | Contextos de Asignatura",
	page_icon="🎯",
	layout="wide",
)


def _inject_styles() -> None:
	st.markdown(
		"""
		<style>
		.hero {
			background: linear-gradient(130deg, #0b3c5d 0%, #173f5f 55%, #20639b 100%);
			color: #ffffff;
			padding: 1.2rem;
			border-radius: 14px;
			border: 1px solid rgba(255,255,255,0.16);
			margin-bottom: 1rem;
		}
		.card {
			background: #f7fbff;
			border: 1px solid #d6e6f5;
			border-radius: 14px;
			padding: 1rem;
			margin-bottom: 0.8rem;
		}
		.tag {
			display: inline-block;
			margin-right: .4rem;
			margin-bottom: .4rem;
			padding: .2rem .6rem;
			border-radius: 999px;
			font-size: 0.82rem;
			background: #eaf3fb;
			border: 1px solid #cbe1f5;
			color: #0b3c5d;
		}
		/* Xuanwu: convertimos el radio del sidebar en botones ergonomicos y modernos. */
		section[data-testid="stSidebar"] .stRadio > div {
			gap: 8px;
		}
		/* Regla 1: ocultar circulos nativos de radio para apariencia tipo boton. */
		section[data-testid="stSidebar"] .stRadio input[type="radio"],
		section[data-testid="stSidebar"] .stRadio label[data-baseweb="radio"] > div:first-child {
			display: none !important;
		}
		/* Regla 2: opciones como botones independientes. */
		section[data-testid="stSidebar"] .stRadio label[data-baseweb="radio"] {
			display: block !important;
			width: 100%;
			padding: 12px 20px !important;
			margin-bottom: 12px !important;
			border-radius: 8px !important;
			border: 1px solid rgba(255, 255, 255, 0.26) !important;
			background-color: rgba(8, 28, 44, 0.38) !important;
			cursor: pointer;
			transition: background-color 0.22s ease, transform 0.18s ease, border-color 0.22s ease;
		}
		section[data-testid="stSidebar"] .stRadio label[data-baseweb="radio"] [data-testid="stMarkdownContainer"] p {
			margin: 0 !important;
			color: #d9e8f6 !important;
			font-weight: 500;
		}
		/* Regla 4: hover ergonomico en opciones inactivas. */
		section[data-testid="stSidebar"] .stRadio label[data-baseweb="radio"]:not(:has(input[type="radio"]:checked)):hover {
			background-color: rgba(29, 74, 109, 0.85) !important;
			border-color: rgba(255, 255, 255, 0.45) !important;
			transform: translateY(-1px);
		}
		/* Regla 3: opcion activa en rojo con texto blanco. */
		section[data-testid="stSidebar"] .stRadio label[data-baseweb="radio"]:has(input[type="radio"]:checked) {
			background-color: #ff4b4b !important;
			border-color: #ff4b4b !important;
		}
		section[data-testid="stSidebar"] .stRadio label[data-baseweb="radio"]:has(input[type="radio"]:checked) [data-testid="stMarkdownContainer"] p {
			color: #ffffff !important;
			font-weight: bold !important;
		}
		section[data-testid="stSidebar"] .stButton button,
		.stButton button {
			padding: 0.6rem 0.9rem !important;
			margin-bottom: 12px !important;
			border-radius: 10px !important;
		}
		</style>
		""",
		unsafe_allow_html=True,
	)


def _init_state() -> None:
	inicializar_si_no_existe()
	if "is_logged" not in st.session_state:
		st.session_state.is_logged = False
	if "usuario_activo" not in st.session_state:
		st.session_state.usuario_activo = ""
	if "username_activo" not in st.session_state:
		st.session_state.username_activo = ""
	if "docente_id" not in st.session_state:
		st.session_state.docente_id = None
	if "usuario_rol" not in st.session_state:
		st.session_state.usuario_rol = ""
	if "unidad_activa_id" not in st.session_state:
		st.session_state.unidad_activa_id = None
	if "borrador_retro" not in st.session_state:
		st.session_state.borrador_retro = ""
	if "evaluacion_activa_id" not in st.session_state:
		st.session_state.evaluacion_activa_id = None
	if "uploader_key" not in st.session_state:
		st.session_state.uploader_key = 1
	if "indicadores_excel_evento" not in st.session_state:
		st.session_state.indicadores_excel_evento = ""
	if "indicadores_excel_procesado" not in st.session_state:
		st.session_state.indicadores_excel_procesado = False


def _slug(texto: str) -> str:
	return "".join(c.lower() if c.isalnum() else "_" for c in texto).strip("_")


def _data_unidad_dir(unidad_id: str) -> Path:
	return ROOT_DIR / "data" / "doctrina" / unidad_id


def _vector_unidad_dir(unidad_id: str) -> Path:
	return ROOT_DIR / "data" / "vector_store" / unidad_id


def _eval_unidad_dir(unidad_id: str) -> Path:
	return ROOT_DIR / "data" / "evaluaciones" / unidad_id


def _limpiar_texto_pdf(texto: str) -> str:
	if not texto:
		return ""
	return " ".join(str(texto).replace("\x00", "").split())


def _leer_texto_docx_desde_bytes(archivo_bytes: bytes) -> str:
	partes: list[str] = []
	try:
		documento = Document(io.BytesIO(archivo_bytes))
		for parrafo in documento.paragraphs:
			texto_parrafo = str(parrafo.text or "").strip()
			if texto_parrafo and not _es_linea_portada_instrumento(texto_parrafo):
				partes.append(texto_parrafo)
	except Exception as e:
		print(f"[DOCX] Error al extraer texto: {e}")
		return "\n".join(partes).strip()
	return "\n".join(partes).strip()


def _leer_texto_pdf_desde_bytes(archivo_bytes: bytes) -> str:
	try:
		pypdf_module = importlib.import_module("pypdf")
		reader = pypdf_module.PdfReader(io.BytesIO(archivo_bytes))
		partes: list[str] = []
		for pagina in reader.pages[:3]:
			partes.append(pagina.extract_text() or "")
		return "\n".join(partes).strip()
	except Exception:
		return ""


def _normalizar_texto_parser(texto: Any) -> str:
	valor = str(texto or "").strip().lower()
	valor = "".join(
		c for c in unicodedata.normalize("NFKD", valor) if not unicodedata.combining(c)
	)
	valor = re.sub(r"\s+", " ", valor)
	return valor


def _clave_orden_pregunta(pregunta_str: str) -> tuple[int, int]:
	roman_map = {
		"I": 1,
		"II": 2,
		"III": 3,
		"IV": 4,
		"V": 5,
		"VI": 6,
		"VII": 7,
		"VIII": 8,
		"IX": 9,
		"X": 10,
	}
	partes = str(pregunta_str).strip().upper().split(".")
	if len(partes) == 2:
		roman, num_str = partes
		val_romano = roman_map.get(roman, 99)
		try:
			val_num = int(num_str)
		except ValueError:
			val_num = 0
		return (val_romano, val_num)
	return (99, 0)


def _normalizar_codigo_req(codigo: str) -> str:
	texto = str(codigo or "").upper()
	texto = texto.replace(",", ".")
	texto = re.sub(r"[^A-Z0-9.]+", "", texto)
	return texto


def _es_req_valido(req: str) -> bool:
	norm = _normalizar_codigo_req(req)
	return bool(re.match(r"^R\d+(?:\.\d+)*$", norm))


def _extraer_preguntas_limpias(texto_celda: str) -> list[str]:
	texto = str(texto_celda or "").upper().replace("\n", " ")
	partes = texto.split(",")
	preguntas: list[str] = []
	vistas: set[str] = set()
	for parte in partes:
		fragmento = parte.strip()
		if not fragmento:
			continue
		match = re.search(r"\b([IVX]+|\d{1,2})\.(\d{1,2})\b", fragmento)
		if not match:
			continue
		izquierda, derecha = match.group(1), match.group(2)
		if int(derecha) <= 0:
			continue
		pregunta = f"{izquierda}.{int(derecha)}"
		if pregunta in vistas:
			continue
		vistas.add(pregunta)
		preguntas.append(pregunta)
	return preguntas


def _es_linea_portada_instrumento(linea: str) -> bool:
	normalizada = _normalizar_texto_parser(linea)
	if not normalizada:
		return True
	frases_bloqueadas = [
		"req",
		"criterios de evaluacion",
		"indicador de evaluacion",
		"aprendizaje esperado por evaluar",
		"tareas asociadas",
		"requerimientos",
		"indicadores de desempeno",
		"unidad de aprendizaje",
		"ejercito de chile",
		"division educacion",
	]
	return any(frase in normalizada for frase in frases_bloqueadas)


def _limpiar_enunciado_pregunta(texto: str) -> str:
	texto_limpio = _limpiar_texto_pdf(texto)
	texto_limpio = texto_limpio.lstrip(". ")
	texto_limpio = re.sub(r"^[\s\.-:;•·,]+", "", texto_limpio)
	texto_limpio = re.sub(r"\s+", " ", texto_limpio).strip()
	return texto_limpio


def _extraer_preguntas_docx_estructuradas(archivo_bytes: bytes) -> list[dict[str, str]]:
	try:
		documento = Document(io.BytesIO(archivo_bytes))
	except Exception as e:
		print(f"[DOCX] Error al abrir el documento: {e}")
		return []

	parrafos: list[str] = []
	for parrafo in documento.paragraphs:
		texto = _limpiar_enunciado_pregunta(parrafo.text)
		if not texto:
			continue
		if _es_linea_portada_instrumento(texto):
			continue
		if re.fullmatch(r"[_\-]{4,}", texto):
			continue
		parrafos.append(texto)

	if not parrafos:
		return []

	patron_item = re.compile(r"^ITEM\s+([IVX]+)\b", re.IGNORECASE)
	patron_pregunta = re.compile(r"^(?:\d+\.\s*)?(¿|[A-Za-zÁÉÍÓÚÑáéíóúñ].*?\?)")
	patron_opcion = re.compile(r"^\s*([a-eA-E])[\)\.\-]\s*(.+)$")
	patron_vf = re.compile(r"^\s*(?:_+\s*)?(.+)$")
	patron_solo_subrayado = re.compile(r"^[_\-\s]+$")
	patron_fin_iii = re.compile(r"^(?:PREGUNTA\s*R\s*4\.1|PREGUNTA\s*R\s*\d|PONDERACI[ÓO]N|NOTA\s+FINAL|PUNTAJE)", re.IGNORECASE)
	patron_inicio_requerimiento = re.compile(r"^\s*(?:\d+\.|[-*•])\s*(.+)$")

	registros: list[dict[str, str]] = []
	seccion_actual: str | None = None
	contador_i = 1
	contador_ii = 1
	contador_iii = 1
	buffer_pregunta_i: list[str] = []
	buffer_opciones_i: list[str] = []
	buffer_iii_contexto: list[str] = []
	buffer_iii_requerimientos: list[str] = []
	capturando_iii = False
	modo_iii_ejercicio = False

	def _cerrar_item_i() -> None:
		nonlocal buffer_pregunta_i, buffer_opciones_i, contador_i
		texto = _limpiar_enunciado_pregunta(" ".join(buffer_pregunta_i + buffer_opciones_i))
		if texto:
			registros.append({"codigo": f"I.{contador_i}", "enunciado": texto, "display": f"Ítem I.{contador_i}: {texto}"})
			contador_i += 1
		buffer_pregunta_i = []
		buffer_opciones_i = []

	def _cerrar_item_ii(texto: str) -> None:
		nonlocal contador_ii
		texto_limpio = _limpiar_enunciado_pregunta(texto)
		if texto_limpio:
			registros.append({"codigo": f"II.{contador_ii}", "enunciado": texto_limpio, "display": f"Ítem II.{contador_ii}: {texto_limpio}"})
			contador_ii += 1

	def _cerrar_item_iii(texto: str) -> None:
		nonlocal contador_iii
		texto_limpio = _limpiar_enunciado_pregunta(texto)
		if texto_limpio:
			registros.append({"codigo": f"III.{contador_iii}", "enunciado": texto_limpio, "display": f"Ítem III.{contador_iii}: {texto_limpio}"})
			contador_iii += 1

	def _agregar_parrafo_fallback(texto: str) -> None:
		nonlocal contador_iii
		texto_limpio = _limpiar_enunciado_pregunta(texto)
		if texto_limpio:
			registros.append({"codigo": f"P{contador_iii}", "enunciado": texto_limpio, "display": f"Ítem P{contador_iii}: {texto_limpio}"})
			contador_iii += 1

	for linea in parrafos:
		if patron_solo_subrayado.match(linea):
			continue

		match_item = patron_item.match(linea)
		if match_item:
			if seccion_actual == "I":
				_cerrar_item_i()
			elif seccion_actual == "III":
				if buffer_iii_contexto:
					_cerrar_item_iii(" ".join(buffer_iii_contexto))
				buffer_iii_contexto = []
				buffer_iii_requerimientos = []
				capturando_iii = False
				modo_iii_ejercicio = False
			seccion_actual = str(match_item.group(1)).upper()
			continue

		if seccion_actual == "I":
			match_opcion = patron_opcion.match(linea)
			if match_opcion and buffer_pregunta_i:
				opcion_texto = f"{match_opcion.group(1).lower()}) {match_opcion.group(2).strip()}"
				buffer_opciones_i.append(opcion_texto)
				continue

			if "¿" in linea or patron_pregunta.match(linea):
				if buffer_pregunta_i:
					_cerrar_item_i()
					buffer_pregunta_i = []
					buffer_opciones_i = []
				pregunta_base = re.sub(r"^\d+\.\s*", "", linea).strip()
				buffer_pregunta_i.append(pregunta_base)
				continue

			if buffer_pregunta_i:
				buffer_pregunta_i.append(linea)
				continue

		if seccion_actual == "II":
			if patron_vf.match(linea) or patron_solo_subrayado.match(linea):
				continue
			if re.search(r"puntaje|instruccion|marque|respuesta", _normalizar_texto_parser(linea)):
				continue
			texto_vf = re.sub(r"^_+\s*", "", linea).lstrip(". ")
			if not texto_vf:
				continue
			_cerrar_item_ii(texto_vf)
			continue

		if seccion_actual == "III":
			if patron_fin_iii.match(linea):
				if buffer_iii_contexto:
					_cerrar_item_iii(" ".join(buffer_iii_contexto))
				break
			texto_norm = _limpiar_enunciado_pregunta(linea)
			if not texto_norm:
				continue
			if "PLANTEAMIENTO" in _normalizar_texto_parser(texto_norm):
				capturando_iii = True
				texto_post = texto_norm.split(":", 1)[1].strip() if ":" in texto_norm else ""
				if texto_post:
					buffer_iii_contexto.append(texto_post)
				continue
			if "EJERCICIO" in _normalizar_texto_parser(texto_norm):
				modo_iii_ejercicio = True
				texto_post = texto_norm.split(":", 1)[1].strip() if ":" in texto_norm else ""
				if texto_post:
					buffer_iii_requerimientos.append(texto_post)
				continue
			if not capturando_iii:
				continue
			if modo_iii_ejercicio:
				match_req = patron_inicio_requerimiento.match(texto_norm)
				if match_req:
					buffer_iii_requerimientos.append(match_req.group(1).strip())
				else:
					buffer_iii_requerimientos.append(texto_norm)
			else:
				buffer_iii_contexto.append(texto_norm)
			continue

	if seccion_actual == "I":
		_cerrar_item_i()

	if seccion_actual == "III":
		if buffer_iii_contexto:
			_cerrar_item_iii(" ".join(buffer_iii_contexto))
		for req in buffer_iii_requerimientos:
			_cerrar_item_iii(req)

	if not registros:
		parrafos_limpios = [
			texto
			for texto in parrafos
			if texto and not _es_linea_portada_instrumento(texto) and not re.fullmatch(r"[_\-\s]+", texto)
		]
		for texto in parrafos_limpios:
			_agregar_parrafo_fallback(texto)

	registros_unicos: list[dict[str, str]] = []
	vistos: set[str] = set()
	for registro in registros:
		codigo = str(registro.get("codigo", "")).strip().upper()
		if not codigo or codigo in vistos:
			continue
		vistos.add(codigo)
		registros_unicos.append(registro)

	return registros_unicos


def _leer_texto_instrumento_desde_path(archivo_path: Path) -> str:
	if not archivo_path.exists():
		return ""
	contenido = archivo_path.read_bytes()
	sufijo = archivo_path.suffix.lower()
	if sufijo == ".docx":
		return _leer_texto_docx_desde_bytes(contenido)
	if sufijo == ".pdf":
		return _leer_texto_pdf_desde_bytes(contenido)
	try:
		return contenido.decode("utf-8")
	except UnicodeDecodeError:
		return contenido.decode("latin-1", errors="ignore")


def _detectar_preguntas_instrumento_detalladas(
	texto: str,
	permitir_secuencial: bool = False,
) -> list[dict[str, str]]:
	lineas_crudas = [re.sub(r"\s+", " ", linea.strip()) for linea in str(texto or "").splitlines()]
	lineas = [linea for linea in lineas_crudas if linea and not _es_linea_portada_instrumento(linea)]
	if not lineas:
		return []

	patron_item = re.compile(r"^\s*ITEM\s+([IVX]+)\b", re.IGNORECASE)
	patron_numero = re.compile(r"^\s*(\d{1,2})\.\s*(.*)$")
	patron_opcion = re.compile(r"^\s*([a-eA-E])[\)\.\-]\s*(.*)$")
	patron_vf = re.compile(r"^\s*(?:_+|V\/?F|F\/?V|VERDADERO\s*\/\s*FALSO|V\s*\/\s*F|F\s*\/\s*V)\b\s*(.*)$", re.IGNORECASE)
	patron_bloque_subrayado = re.compile(r"^[_\-\s]{4,}$")
	patron_inicio_enunciado = re.compile(r"(?:^|\s)(?:¿|[A-ZÁÉÍÓÚÑ].*\?)")
	patron_texto_interrogativo = re.compile(r"^(cual|cuál|como|cómo|que|qué|donde|dónde|cuando|cuándo|quien|quién|por que|por qué)\b", re.IGNORECASE)
	patron_respuesta = re.compile(r"^\s*(?:[vVfF]|verdadero|falso|si|sí|no)\s*$", re.IGNORECASE)

	def _es_basura(linea: str) -> bool:
		if not linea:
			return True
		if patron_bloque_subrayado.match(linea):
			return True
		normalizada = _normalizar_texto_parser(linea)
		return not any(ch.isalnum() for ch in normalizada)

	def _es_inicio_pregunta(linea: str) -> bool:
		if patron_numero.match(linea):
			return True
		if patron_vf.match(linea):
			return True
		if "¿" in linea or "?" in linea:
			return True
		if patron_texto_interrogativo.match(linea):
			return True
		if patron_inicio_enunciado.search(linea) and not patron_opcion.match(linea):
			return True
		return False

	def _limpiar_bloque_enunciado(bloque: list[str]) -> str:
		texto_bloque = " ".join(bloque)
		texto_bloque = _limpiar_enunciado_pregunta(texto_bloque)
		texto_bloque = re.sub(r"\s{2,}", " ", texto_bloque).strip()
		return texto_bloque

	def _fallback_por_parrafos(parrafos: list[str]) -> list[dict[str, str]]:
		registros: list[dict[str, str]] = []
		secuencia = 1
		for parrafo in parrafos:
			texto_parrafo = _limpiar_enunciado_pregunta(parrafo)
			if not texto_parrafo or _es_basura(texto_parrafo):
				continue
			codigo = f"P{secuencia}"
			registros.append(
				{
					"codigo": codigo,
					"enunciado": texto_parrafo,
					"display": f"Ítem {codigo}: {texto_parrafo}",
				}
			)
			secuencia += 1
		return registros

	secciones: dict[str, list[str]] = {"I": [], "II": [], "III": []}
	seccion_actual: str | None = None
	for linea in lineas:
		match_item = patron_item.match(linea)
		if match_item:
			seccion_actual = str(match_item.group(1)).upper()
			continue
		if seccion_actual in secciones:
			secciones[seccion_actual].append(linea)

	def _es_item_i(linea: str) -> bool:
		return _es_inicio_pregunta(linea) and not patron_respuesta.match(linea) and not _es_basura(linea)

	def _parsear_item_i(lineas_item: list[str]) -> list[dict[str, str]]:
		registros: list[dict[str, str]] = []
		pregunta_actual: list[str] = []
		opciones_actuales: list[str] = []
		numero = 1

		def _cerrar() -> None:
			nonlocal pregunta_actual, opciones_actuales, numero
			texto = _limpiar_bloque_enunciado(pregunta_actual + opciones_actuales)
			if texto:
				registros.append(
					{
						"codigo": f"I.{numero}",
						"enunciado": texto,
						"display": f"Ítem I.{numero}: {texto}",
					}
				)
				numero += 1
			pregunta_actual = []
			opciones_actuales = []

		for linea in lineas_item:
			if _es_basura(linea):
				continue
			match_num = patron_numero.match(linea)
			match_opcion = patron_opcion.match(linea)
			if match_num and (_es_item_i(linea) or pregunta_actual):
				if pregunta_actual or opciones_actuales:
					_cerrar()
				pregunta_actual = [match_num.group(2).strip() or linea]
				opciones_actuales = []
				continue
			if match_opcion and pregunta_actual:
				opciones_actuales.append(f"{match_opcion.group(1).lower()}) {match_opcion.group(2).strip()}")
				continue
			if patron_vf.match(linea):
				continue
			if not pregunta_actual:
				if "¿" in linea or patron_texto_interrogativo.match(linea):
					pregunta_actual = [linea]
				continue
			if pregunta_actual and not match_opcion:
				pregunta_actual.append(linea)
		_cerrar()
		return registros

	def _parsear_item_ii(lineas_item: list[str]) -> list[dict[str, str]]:
		registros: list[dict[str, str]] = []
		buffer: list[str] = []
		numero = 1
		for linea in lineas_item:
			if _es_basura(linea) or patron_respuesta.match(linea):
				continue
			if patron_vf.match(linea):
				continue
			if patron_bloque_subrayado.match(linea):
				if buffer:
					texto = _limpiar_bloque_enunciado(buffer)
					if texto:
						registros.append({"codigo": f"II.{numero}", "enunciado": texto, "display": f"Ítem II.{numero}: {texto}"})
						numero += 1
					buffer = []
				continue
			texto_linea = linea
			texto_linea = re.sub(r"^_+\s*", "", texto_linea)
			texto_linea = texto_linea.lstrip(". ")
			if texto_linea:
				buffer.append(texto_linea)
				if len(buffer) >= 1:
					texto = _limpiar_bloque_enunciado(buffer)
					if texto:
						registros.append({"codigo": f"II.{numero}", "enunciado": texto, "display": f"Ítem II.{numero}: {texto}"})
						numero += 1
						buffer = []
						continue
		if buffer:
			texto = _limpiar_bloque_enunciado(buffer)
			if texto:
				registros.append({"codigo": f"II.{numero}", "enunciado": texto, "display": f"Ítem II.{numero}: {texto}"})
		return registros

	def _parsear_item_iii(lineas_item: list[str]) -> list[dict[str, str]]:
		registros: list[dict[str, str]] = []
		contexto: list[str] = []
		requerimientos: list[str] = []
		numero = 1
		for linea in lineas_item:
			if _es_basura(linea) or patron_respuesta.match(linea):
				continue
			match_num = patron_numero.match(linea)
			if match_num:
				requerimiento = match_num.group(2).strip() or linea
				requerimiento = _limpiar_enunciado_pregunta(requerimiento)
				if requerimiento:
					requerimientos.append(requerimiento)
				continue
			if not requerimientos:
				contexto.append(linea)
			else:
				requerimientos[-1] = f"{requerimientos[-1]} {linea}".strip()

		contexto_limpio = _limpiar_bloque_enunciado(contexto)
		if not requerimientos and contexto_limpio:
			requerimientos = [contexto_limpio]

		for requerimiento in requerimientos:
			enunciado = requerimiento
			if contexto_limpio and requerimiento != contexto_limpio:
				enunciado = f"{contexto_limpio} {requerimiento}".strip()
			registros.append({"codigo": f"III.{numero}", "enunciado": enunciado, "display": f"Ítem III.{numero}: {enunciado}"})
			numero += 1
		return registros

	registros: list[dict[str, str]] = []
	registros.extend(_parsear_item_i(secciones["I"]))
	registros.extend(_parsear_item_ii(secciones["II"]))
	registros.extend(_parsear_item_iii(secciones["III"]))

	if not registros:
		parrafos = [
			_limpiar_enunciado_pregunta(linea)
			for linea in lineas
			if linea and not _es_linea_portada_instrumento(linea) and not _es_basura(linea)
		]
		return _fallback_por_parrafos(parrafos)

	registros_unicos: list[dict[str, str]] = []
	vistos_codigos: set[str] = set()
	for registro in registros:
		codigo = str(registro.get("codigo", "")).strip().upper()
		if not codigo or codigo in vistos_codigos:
			continue
		vistos_codigos.add(codigo)
		registros_unicos.append(registro)

	return registros_unicos


def _detectar_preguntas_instrumento(texto: str, permitir_secuencial: bool = False) -> list[str]:
	return [item["codigo"] for item in _detectar_preguntas_instrumento_detalladas(texto, permitir_secuencial)]


def _resolver_indicador_asociado(req: str, indicador_texto: str, indicadores: list[str]) -> str:
	if not indicadores:
		return _normalizar_codigo_req(req) or "R4.1.1"

	req_norm = _normalizar_codigo_req(req)
	ind_norm = _normalizar_codigo_req(indicador_texto)
	indicadores_norm = [str(i).strip().upper() for i in indicadores]

	for codigo in indicadores_norm:
		if ind_norm and (codigo == ind_norm or codigo.startswith(ind_norm) or ind_norm.startswith(codigo)):
			return codigo

	for codigo in indicadores_norm:
		if req_norm and (codigo == req_norm or codigo.startswith(req_norm)):
			return codigo

	return indicadores_norm[0]


def _normalizar_lista_indicadores(indicadores_texto: str) -> str:
	tokens = re.split(r"[,;\n]+", str(indicadores_texto or ""))
	limpios: list[str] = []
	vistos: set[str] = set()
	for token in tokens:
		codigo = _normalizar_codigo_req(token)
		if not codigo:
			continue
		if codigo in vistos:
			continue
		vistos.add(codigo)
		limpios.append(codigo)
	return ", ".join(limpios)


def _parsear_indicadores_asociados(indicadores_texto: str) -> list[str]:
	tokens = re.split(r"[,;\n]+", str(indicadores_texto or ""))
	resultados: list[str] = []
	vistos: set[str] = set()
	for token in tokens:
		codigo = _normalizar_codigo_req(token)
		if not codigo:
			continue
		if codigo in vistos:
			continue
		vistos.add(codigo)
		resultados.append(codigo)
	return resultados


def _consolidar_pregunta_indicadores(filas: list[dict[str, Any]]) -> list[dict[str, Any]]:
	acumulado: dict[str, dict[str, Any]] = {}
	for fila in filas:
		numero = str(fila.get("numero_pregunta", "")).strip().upper()
		if not numero:
			continue
		try:
			puntos = float(fila.get("puntos_maximos", 0.0))
		except (TypeError, ValueError):
			puntos = 0.0
		indicadores = _parsear_indicadores_asociados(str(fila.get("indicadores_asociados", "")))
		if not indicadores:
			continue

		if numero not in acumulado:
			acumulado[numero] = {"puntos_maximos": puntos, "indicadores": []}
		elif puntos > 0:
			# Conservamos el mayor puntaje si aparece repetido para la misma pregunta.
			acumulado[numero]["puntos_maximos"] = max(float(acumulado[numero]["puntos_maximos"]), puntos)

		for ind in indicadores:
			if ind not in acumulado[numero]["indicadores"]:
				acumulado[numero]["indicadores"].append(ind)

	salida: list[dict[str, Any]] = []
	for numero in sorted(acumulado.keys(), key=_clave_orden_pregunta):
		info = acumulado[numero]
		salida.append(
			{
				"numero_pregunta": numero,
				"puntos_maximos": float(info["puntos_maximos"]),
				"indicadores_asociados": ", ".join(info["indicadores"]),
			}
		)
	return salida


def _extraer_mapeo_docx_tabla_tributacion(
	archivo_bytes: bytes,
	indicadores: list[str],
) -> list[dict[str, Any]]:
	try:
		doc = Document(io.BytesIO(archivo_bytes))
	except Exception:
		return []

	for table in doc.tables:
		filas: list[list[str]] = []
		for row in table.rows:
			fila_textos = [cell.text.strip().replace("\n", " ") for cell in row.cells]
			fila_textos = [re.sub(r"\s+", " ", texto).strip() for texto in fila_textos]
			if any(fila_textos):
				filas.append(fila_textos)

		header_idx = -1
		for i, fila in enumerate(filas):
			norm_fila = [_normalizar_texto_parser(celda) for celda in fila]
			tiene_req = any("req" in celda for celda in norm_fila)
			tiene_pregunta = any("pregunta" in celda for celda in norm_fila)
			if tiene_req and tiene_pregunta:
				header_idx = i
				break

		if header_idx < 0:
			continue

		propuesta_raw: list[dict[str, Any]] = []
		ultimo_req_valido = ""
		for fila in filas[header_idx + 1:]:
			req_celda = fila[0] if len(fila) > 0 else ""
			req = str(req_celda or "").strip()
			if _es_req_valido(req):
				ultimo_req_valido = req
			elif ultimo_req_valido:
				req = ultimo_req_valido
			else:
				continue

			preguntas_txt = fila[-1] if fila else ""
			preguntas = _extraer_preguntas_limpias(preguntas_txt)
			if not preguntas:
				continue

			indicador_txt = fila[4] if len(fila) > 4 else ""
			indicador_asociado = _resolver_indicador_asociado(req, indicador_txt, indicadores)

			for pregunta in preguntas:
				propuesta_raw.append(
					{
						"numero_pregunta": pregunta,
						"puntos_maximos": 1.0,
						"indicadores_asociados": indicador_asociado,
					}
				)

		if propuesta_raw:
			return _consolidar_pregunta_indicadores(propuesta_raw)

	return []


def _extraer_mapeo_pdf_tabla_tributacion(
	archivo_bytes: bytes,
	indicadores: list[str],
) -> list[dict[str, Any]]:
	texto = _leer_texto_pdf_desde_bytes(archivo_bytes)
	norm = _normalizar_texto_parser(texto)
	requeridos = ["req", "criterios de evaluacion", "indicador de evaluacion", "pregunta"]
	if not all(token in norm for token in requeridos):
		return []

	propuesta_raw: list[dict[str, Any]] = []
	for linea in texto.splitlines():
		linea_limpia = linea.strip()
		if not linea_limpia:
			continue

		match_req = re.match(r"^\s*(R\s*\d+(?:[\.,]\d+)*)", linea_limpia, flags=re.IGNORECASE)
		if not match_req:
			continue

		req = match_req.group(1)
		if not _es_req_valido(req):
			continue

		preguntas = _extraer_preguntas_limpias(linea_limpia)
		if not preguntas:
			continue

		indicador_asociado = _resolver_indicador_asociado(req, linea_limpia, indicadores)
		for pregunta in preguntas:
			propuesta_raw.append(
				{
					"numero_pregunta": pregunta,
					"puntos_maximos": 1.0,
					"indicadores_asociados": indicador_asociado,
				}
			)

	return _consolidar_pregunta_indicadores(propuesta_raw)


def _normalizar_columna_excel(nombre_columna: Any) -> str:
	texto = str(nombre_columna).strip().lower()
	texto = "".join(
		c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
	)
	texto = re.sub(r"[^a-z0-9]+", " ", texto).strip()
	return re.sub(r"\s+", " ", texto)


def _mapear_columnas_indicadores_excel(df: pd.DataFrame) -> pd.DataFrame | None:
	alias_codigo = {
		"codigo",
		"codigo completo",
		"codigo indicador",
		"codigo del indicador",
		"codigo completo indicador",
	}
	alias_descripcion = {
		"descripcion",
		"descripcion indicador",
		"descripcion del indicador",
		"detalle",
		"detalle indicador",
	}

	candidatas_codigo: list[str] = []
	candidatas_descripcion: list[str] = []

	for col in df.columns:
		norm = _normalizar_columna_excel(col)
		es_codigo = norm in alias_codigo or (
			"codigo" in norm and ("completo" in norm or "indicador" in norm or norm == "codigo")
		)
		es_descripcion = norm in alias_descripcion or "descripcion" in norm or "detalle" in norm

		if es_codigo:
			candidatas_codigo.append(str(col))
		if es_descripcion:
			candidatas_descripcion.append(str(col))

	if len(candidatas_codigo) != 1 or len(candidatas_descripcion) != 1:
		return None

	col_codigo = candidatas_codigo[0]
	col_descripcion = candidatas_descripcion[0]
	if col_codigo == col_descripcion:
		return None

	return df.rename(columns={col_codigo: "Codigo_Completo", col_descripcion: "Descripcion"})


def _unidad_activa() -> dict[str, Any] | None:
	uid = st.session_state.unidad_activa_id
	if not uid:
		return None
	try:
		return obtener_unidad_por_id(int(uid))
	except (TypeError, ValueError):
		return None


def _render_login() -> None:
	st.markdown(
		"""
		<div class="hero">
			<h2>🛡️ Acceso Institucional ACAPOMIL</h2>
			<p>Autenticacion requerida. El sistema se habilita solo con sesion valida.</p>
		</div>
		""",
		unsafe_allow_html=True,
	)

	col1, col2 = st.columns([1.2, 1])
	with col1:
		with st.container(border=True):
			tab_login, tab_registro = st.tabs(["🔐 Iniciar Sesion", "📝 Registrarse"])

			with tab_login:
				st.subheader("Acceso docente")
				usuario = st.text_input(
					"Usuario",
					key="login_usuario",
					placeholder="Ingrese su nombre de usuario o correo",
				)
				clave = st.text_input("Clave institucional", type="password", key="login_clave")

				if st.button("Ingresar al sistema", use_container_width=True, type="primary"):
					# Xuanwu: autenticamos contra SQLite para persistencia real.
					credenciales = obtener_usuario_por_credenciales(usuario.strip(), clave)
					if credenciales:
						st.session_state.is_logged = True
						st.session_state.usuario_activo = credenciales["nombre_completo"]
						st.session_state.username_activo = credenciales["username"]
						st.session_state.docente_id = int(credenciales["id"])
						st.session_state.usuario_rol = str(credenciales.get("rol", "docente"))
						st.session_state.unidad_activa_id = None
						st.rerun()
					st.error("Credenciales invalidas para acceso docente.")

			with tab_registro:
				st.subheader("Registro de nuevo docente")
				with st.form("registro_docente_form", clear_on_submit=True):
					nuevo_username = st.text_input("Username", placeholder="usuario.docente")
					nuevo_nombre_completo = st.text_input("Nombre completo", placeholder="Nombre Apellido")
					nueva_clave = st.text_input("Clave", type="password")
					registro_ok = st.form_submit_button("Registrar docente", use_container_width=True)

				if registro_ok:
					username = nuevo_username.strip()
					nombre_completo = nuevo_nombre_completo.strip()
					if not username or not nombre_completo or not nueva_clave.strip():
						st.error("Completa username, nombre completo y clave para registrarte.")
					elif obtener_usuario_por_username(username):
						st.warning("Ese usuario ya existe. Usa otro nombre o inicia sesion.")
					else:
						# Xuanwu: registro docente persistente en SQLite.
						crear_usuario(username, nombre_completo, nueva_clave, "docente")
						st.success("Registro exitoso. Ya puedes iniciar sesion en la pestana de acceso.")

	with col2:
		with st.container(border=True):
			st.markdown("### 🎯 Perfil habilitado")
			st.markdown("- Docente por defecto: **123** (123)")
			st.markdown("- Admin por defecto: **admin**")
			st.markdown("- Clave inicial docente por defecto: **123**")
			st.markdown("- Clave inicial admin por defecto: **admin123**")
			st.markdown("- Ambito: Evaluacion por competencias")
			st.markdown("- Seguridad: Sesion obligatoria antes de operar")


def _render_top_context() -> None:
	if st.session_state.usuario_rol == "admin":
		st.markdown(
			"""
			<div class="hero">
				<h2>🛡️ Panel de Administracion ACAPOMIL</h2>
				<p>Control centralizado de matricula semestral por CRIM y Mencion.</p>
			</div>
			""",
			unsafe_allow_html=True,
		)
		st.markdown("**Panel de Control:** admin (Administrador)")
		return

	unidad = _unidad_activa()
	st.markdown(
		"""
		<div class="hero">
			<h2>📚 Contextos de Asignatura</h2>
			<p>Operacion tactica de retroalimentacion RAG por unidad de aprendizaje.</p>
		</div>
		""",
		unsafe_allow_html=True,
	)
	st.markdown(f"**Docente:** {st.session_state.usuario_activo}")

	if unidad:
		st.markdown(
			f"<span class='tag'>Unidad activa: {unidad['nombre']}</span>"
			f"<span class='tag'>Año: {unidad['anio']}</span>"
			f"<span class='tag'>Mencion: {unidad['mencion']}</span>",
			unsafe_allow_html=True,
		)
	else:
		st.warning("No hay Unidad Activa seleccionada. Define una para habilitar el resto del flujo.")


def _menu_sidebar() -> str:
	st.sidebar.title("🧭 Modulos")
	if st.session_state.usuario_rol == "admin":
		st.sidebar.success("Perfil: Administrador")
	else:
		unidad = _unidad_activa()
		if unidad:
			st.sidebar.success(f"Unidad activa: {unidad['nombre']}")
		else:
			st.sidebar.info("Sin unidad activa")

	if st.sidebar.button("Cerrar sesion", use_container_width=True):
		st.session_state.is_logged = False
		st.session_state.usuario_activo = ""
		st.session_state.username_activo = ""
		st.session_state.docente_id = None
		st.session_state.usuario_rol = ""
		st.session_state.unidad_activa_id = None
		st.rerun()

	if st.session_state.usuario_rol == "admin":
		return st.sidebar.radio(
			"Navegacion",
			[
				"1) Control de Matricula",
				"2) Gestion de Cuentas Docentes",
			],
		)

	return st.sidebar.radio(
		"Navegacion",
		[
			"1) Gestion de Unidades",
			"2) Diseño de Evaluación",
			"3) Bibliografia y Vectorizacion",
			"4) Trazabilidad RAG",
			"5) Reportes Jefatura",
		],
	)


def _vista_admin_matricula() -> None:
	st.subheader("🗂️ Control de Matricula Semestral")
	with st.container(border=True):
		st.info("Instrucciones de uso: administra altas y bajas de alumnos por CRIM y Mencion. Este panel es exclusivo para Administrador.")

	with st.container(border=True):
		st.markdown("### ➕ Matricular alumno")
		c1, c2 = st.columns(2)
		with c1:
			nombre = st.text_input("Nombre completo", key="admin_nombre")
			rut = st.text_input("RUT", placeholder="12.345.678-9", key="admin_rut")
		with c2:
			anio = st.selectbox("Año", ANIOS_CRIM, key="admin_anio")
			mencion = st.selectbox("Mencion", MENCIONES, key="admin_mencion")

		if st.button("Guardar matricula", use_container_width=True, type="primary"):
			rut_ingresado = rut.strip()
			if not nombre.strip() or not rut_ingresado:
				st.error("Nombre y RUT son obligatorios.")
			elif not validar_rut_chileno(rut_ingresado):
				st.error("❌ El RUT ingresado es inválido según el algoritmo estándar chileno. Por favor, verifíquelo.")
			else:
				rut_formateado = formatear_rut_chileno(rut_ingresado)
				if buscar_alumno_por_rut_anio_mencion(rut=rut_formateado, anio=anio, mencion=mencion):
					st.warning("Ese alumno ya esta matriculado para ese CRIM y Mencion.")
				else:
					# Xuanwu: matricula centralizada persistente para consumo automatico docente.
					crear_alumno(rut=rut_formateado, nombre=nombre.strip(), anio=anio, mencion=mencion)
					st.success("Alumno matriculado correctamente.")
					st.rerun()

	with st.container(border=True):
		st.markdown("### 🧾 Matricula registrada")
		filtro_anio = st.selectbox("Filtrar por Año", ["Todos"] + ANIOS_CRIM, key="admin_filtro_anio")
		filtro_mencion = st.selectbox("Filtrar por Mencion", ["Todas"] + MENCIONES, key="admin_filtro_mencion")
		anio_sel = None if filtro_anio == "Todos" else filtro_anio
		mencion_sel = None if filtro_mencion == "Todas" else filtro_mencion
		alumnos = listar_alumnos(anio=anio_sel, mencion=mencion_sel)

		if not alumnos:
			st.info("No hay alumnos matriculados para el filtro seleccionado.")
			return

		tabla = pd.DataFrame(alumnos)
		tabla = tabla[["rut", "nombre", "anio", "mencion"]]
		st.dataframe(tabla, use_container_width=True)

		mapa = {f"{a['nombre']} ({a['rut']}) | {a['anio']} | {a['mencion']}": int(a["id"]) for a in alumnos}
		obj = st.selectbox("Alumno a eliminar", list(mapa.keys()))
		if st.button("Eliminar matricula", use_container_width=True):
			eliminar_alumno(mapa[obj])
			st.success("Matricula eliminada.")
			st.rerun()


def _vista_admin_cuentas() -> None:
	st.subheader("👤 Gestion de Cuentas Docentes")
	with st.container(border=True):
		st.info("Instrucciones de uso: crea, actualiza o elimina cuentas. Por seguridad, la cuenta admin no puede eliminarse.")

	usuarios = obtener_usuarios()
	with st.container(border=True):
		st.markdown("### 📋 Cuentas registradas")
		if not usuarios:
			st.info("No hay usuarios registrados.")
		else:
			tabla = pd.DataFrame(usuarios)
			tabla = tabla[["nombre_completo", "username", "rol"]]
			st.dataframe(tabla, use_container_width=True)

	with st.container(border=True):
		st.markdown("### ➕ Crear cuenta docente")
		col_crear_1, col_crear_2 = st.columns(2)
		with col_crear_1:
			new_username = st.text_input("Username")
			new_nombre = st.text_input("Nombre Completo (con Grado Militar)")
		with col_crear_2:
			new_clave = st.text_input("Clave", type="password")
			new_rol = st.selectbox("Asignar Rol", ["docente", "admin"])

		if st.button("Registrar cuenta", use_container_width=True):
			if not new_username or not new_clave or not new_nombre:
				st.warning("⚠️ Por favor, rellene todos los campos antes de registrar.")
			else:
				try:
					# Intentamos la insercion en la base de datos
					crear_usuario(
						username=new_username.strip(),
						nombre_completo=new_nombre.strip(),
						clave=new_clave.strip(),
						rol=new_rol,
					)
					st.success("🎉 ¡Cuenta registrada con éxito!")
					st.rerun()
				except Exception as e:
					st.error(f"❌ Error crítico en el backend: {str(e)}")

	with st.container(border=True):
		st.markdown("### ✏️ Editar o eliminar cuenta")
		if not usuarios:
			st.info("Sin cuentas para administrar.")
			return

		usernames = [u["username"] for u in usuarios]
		username_objetivo = st.selectbox("Cuenta objetivo", usernames, key="admin_user_target")
		objetivo = next(u for u in usuarios if u["username"] == username_objetivo)

		col1, col2 = st.columns(2)
		with col1:
			ed_username = st.text_input("Nuevo username", value=objetivo["username"], key="admin_user_edit_username")
			ed_nombre_completo = st.text_input("Nuevo nombre completo", value=objetivo["nombre_completo"], key="admin_user_edit_fullname")
		with col2:
			ed_clave = st.text_input("Nueva clave", type="password", key="admin_user_edit_pwd", placeholder="Dejar vacio para mantener")
			opciones_rol = ["docente", "admin"]
			rol_actual = str(objetivo.get("rol", "docente")).strip().lower()
			rol_index = opciones_rol.index(rol_actual) if rol_actual in opciones_rol else 0
			ed_rol = st.selectbox(
				"Rol",
				opciones_rol,
				index=rol_index,
				key=f"admin_user_edit_role_{username_objetivo}",
			)

		confirmar = st.checkbox("Confirmar eliminacion", key="admin_user_delete_confirm")

		act_col1, act_col2 = st.columns(2)
		with act_col1:
			if st.button("Guardar cambios cuenta", use_container_width=True):
				username_final = ed_username.strip() or objetivo["username"]
				nombre_completo_final = ed_nombre_completo.strip() or objetivo["nombre_completo"]
				existente = obtener_usuario_por_username(username_final)
				if username_final != objetivo["username"] and existente:
					st.warning("Ya existe otra cuenta con ese nombre.")
				elif objetivo["username"] == "admin" and ed_rol != "admin":
					st.error("La cuenta admin no puede cambiar su rol a docente.")
				else:
					try:
						actualizar_usuario(
							usuario_id=int(objetivo["id"]),
							username=username_final,
							nombre_completo=nombre_completo_final,
							clave=ed_clave if ed_clave.strip() else None,
							rol=ed_rol,
						)
						st.success("Cuenta actualizada.")
						st.rerun()
					except ValueError as exc:
						st.error(str(exc))

		with act_col2:
			if st.button("Eliminar cuenta", use_container_width=True, type="secondary"):
				if not confirmar:
					st.warning("Debes confirmar antes de eliminar.")
				elif username_objetivo == "admin" or int(objetivo["id"]) == int(st.session_state.docente_id or -1):
					st.error("No puedes eliminar la cuenta administradora activa.")
				else:
					try:
						# Xuanwu: eliminacion robusta por username unico seleccionado en el selectbox.
						eliminar_usuario_por_username(username_objetivo)
						st.success("Cuenta eliminada.")
						st.rerun()
					except ValueError as exc:
						st.error(str(exc))
					except sqlite3.IntegrityError:
						st.error("No se puede eliminar: Este docente tiene unidades de aprendizaje asignadas en el sistema.")
					except Exception:
						st.error("Error al intentar eliminar esta cuenta. Verifica las dependencias en la base de datos.")


def _vista_unidades() -> None:
	st.subheader("🏛️ Gestion de Unidades de Aprendizaje")
	with st.container(border=True):
		st.info("Instrucciones de uso: crea una unidad, activala y luego edita o elimina segun necesidad academica.")

	docente_id = st.session_state.docente_id
	if docente_id is None:
		st.error("Sesion docente no valida. Vuelve a iniciar sesion.")
		return

	unidades = listar_unidades(docente_id=docente_id)

	c1, c2 = st.columns([1.2, 1])
	with c1:
		with st.container(border=True):
			st.markdown("### ➕ Crear unidad")
			with st.form("crear_unidad_form", clear_on_submit=True):
				nombre = st.text_input("Nombre de unidad", placeholder="Bases de Datos")
				anio = st.selectbox("Año", ANIOS_CRIM)
				mencion = st.selectbox("Mencion", MENCIONES)
				submitted = st.form_submit_button("Crear unidad", use_container_width=True)

			if submitted:
				if not nombre.strip():
					st.error("Ingresa un nombre valido para la unidad.")
				else:
					# Xuanwu: alta de unidad persistente en base relacional.
					unidad_id = crear_unidad(nombre.strip(), anio, mencion, docente_id)
					if not st.session_state.unidad_activa_id:
						st.session_state.unidad_activa_id = unidad_id
					# Xuanwu: preparamos carpetas por unidad para doctrina y vector store.
					_data_unidad_dir(str(unidad_id)).mkdir(parents=True, exist_ok=True)
					_vector_unidad_dir(str(unidad_id)).mkdir(parents=True, exist_ok=True)
					st.success("Unidad creada y disponible para seleccion.")
					st.rerun()

	with c2:
		with st.container(border=True):
			st.markdown("### 🎛️ Seleccionar unidad activa")
			if not unidades:
				st.info("Aun no existen unidades.")
			else:
				mapa = {
					f"{u['nombre']} | {u['anio']} | {u['mencion']}": u["id"]
					for u in unidades
				}
				opciones = list(mapa.keys())
				idx = 0
				if st.session_state.unidad_activa_id is not None:
					for i, key in enumerate(opciones):
						if mapa[key] == st.session_state.unidad_activa_id:
							idx = i
							break
				seleccion = st.selectbox("Unidad", opciones, index=idx)
				if st.button("Activar unidad", use_container_width=True):
					st.session_state.unidad_activa_id = mapa[seleccion]
					st.success("Unidad activa actualizada.")
					# Xuanwu: refresco explicito para actualizar cabecera y sidebar de inmediato.
					st.rerun()

	with st.container(border=True):
		st.markdown("### 🛠️ Editar / Eliminar unidad")
		if not unidades:
			st.info("Sin unidades para administrar.")
			return

		mapa = {f"{u['nombre']} ({u['anio']})": u for u in unidades}
		pick = st.selectbox("Unidad objetivo", list(mapa.keys()), key="editar_unidad_pick")
		target = mapa[pick]

		ncol1, ncol2, ncol3 = st.columns(3)
		with ncol1:
			nuevo_nombre = st.text_input("Nuevo nombre", value=target["nombre"], key="ed_nombre")
		with ncol2:
			nuevo_anio = st.selectbox(
				"Nuevo año",
				ANIOS_CRIM,
				index=ANIOS_CRIM.index(target["anio"]),
				key="ed_anio",
			)
		with ncol3:
			nueva_mencion = st.selectbox(
				"Nueva mencion",
				MENCIONES,
				index=MENCIONES.index(target["mencion"]),
				key="ed_mencion",
			)

		a, b = st.columns(2)
		with a:
			if st.button("Guardar cambios", use_container_width=True):
				actualizar_unidad(
					unidad_id=int(target["id"]),
					nombre=nuevo_nombre.strip() or target["nombre"],
					anio=nuevo_anio,
					mencion=nueva_mencion,
				)
				st.success("Unidad actualizada.")
				st.rerun()
		with b:
			confirmar = st.checkbox("Confirmar eliminacion", key="confirm_del_unit")
			if st.button("Eliminar unidad", use_container_width=True, type="secondary"):
				if not confirmar:
					st.warning("Marca la confirmacion para eliminar.")
				else:
					uid = int(target["id"])
					eliminar_unidad(uid)
					shutil.rmtree(_data_unidad_dir(str(uid)), ignore_errors=True)
					shutil.rmtree(_vector_unidad_dir(str(uid)), ignore_errors=True)
					if st.session_state.unidad_activa_id == uid:
						restantes = listar_unidades(docente_id=docente_id)
						st.session_state.unidad_activa_id = restantes[0]["id"] if restantes else None
					st.success("Unidad eliminada.")
					st.rerun()


def _vista_diseno_evaluacion() -> None:
	st.subheader("🧩 Diseño y Mapeo de Evaluaciones")
	with st.container(border=True):
		st.info("Instrucciones de uso: define indicadores, carga el instrumento y valida su matriz de tributacion pregunta-indicador antes de evaluar.")

	unidad = _unidad_activa()
	if not unidad:
		st.info("Activa una unidad para habilitar el diseño de evaluacion.")
		return

	unidad_id = int(unidad["id"])

	with st.container(border=True):
		st.markdown("### Fase 1: Indicadores de la Unidad Activa")
		cargador_col, info_col = st.columns(2)
		with cargador_col:
			archivo_indicadores = st.file_uploader(
				"Subir Matriz de Indicadores (.xlsx / .pdf)",
				type=["xlsx", "pdf"],
				key=f"excel_uploader_{st.session_state.uploader_key}",
			)
		with info_col:
			st.caption("Excel: columnas Codigo_Completo y Descripcion. PDF: texto con codigo y descripcion.")

		if archivo_indicadores is not None:
			evento_actual = f"{unidad_id}:{archivo_indicadores.name}:{archivo_indicadores.size}:{st.session_state.uploader_key}"
			if st.session_state.indicadores_excel_evento != evento_actual:
				st.session_state.indicadores_excel_evento = evento_actual
				st.session_state.indicadores_excel_procesado = False

			if not st.session_state.indicadores_excel_procesado:
				st.session_state.indicadores_excel_procesado = True
				if archivo_indicadores.name.lower().endswith(".pdf"):
					try:
						indicadores_pdf = extraer_indicadores_desde_pdf(archivo_indicadores.getvalue())
					except Exception as exc:
						st.error(f"❌ Error al leer el PDF: {exc}")
					else:
						indicadores_limpios: list[dict[str, str]] = []
						vistos: set[str] = set()
						for fila in indicadores_pdf:
							codigo = str(fila.get("codigo", "")).strip().upper()
							descripcion = str(fila.get("descripcion", "")).strip()
							if not codigo or codigo in vistos:
								continue
							vistos.add(codigo)
							indicadores_limpios.append({"codigo": codigo, "descripcion": descripcion})

						if not indicadores_limpios:
							st.error("No se encontraron indicadores válidos en el PDF.")
						else:
							reemplazar_indicadores_unidad(unidad_id, indicadores_limpios)
							st.success(f"🎉 ¡Se extrajeron y guardaron {len(indicadores_limpios)} indicadores desde el PDF!")
							st.session_state.uploader_key += 1
							st.session_state.indicadores_excel_evento = ""
							st.session_state.indicadores_excel_procesado = False
							st.rerun()
				else:
					try:
						df_indicadores = pd.read_excel(archivo_indicadores)
					except Exception as exc:
						st.error(f"❌ Error al leer el archivo Excel: {exc}")
					else:
						df_mapeado = _mapear_columnas_indicadores_excel(df_indicadores)
						if df_mapeado is None:
							st.error("❌ Formato de Excel inválido. Debe contener las columnas 'Codigo_Completo' y 'Descripcion'.")
						else:
							cargar_indicadores_desde_df(unidad_id, df_mapeado)
							st.success(f"🎉 ¡Se cargaron con éxito {len(df_mapeado)} indicadores para esta Unidad de Aprendizaje!")
							st.session_state.uploader_key += 1
							st.session_state.indicadores_excel_evento = ""
							st.session_state.indicadores_excel_procesado = False
							st.rerun()
		else:
			indicadores_db = listar_indicadores_unidad(unidad_id)
			base_rows = [
				{"codigo": row["codigo"], "descripcion": row.get("descripcion", "") or ""}
				for row in indicadores_db
			] or [{"codigo": "R4.1.1", "descripcion": ""}]

			editor_indicadores = st.data_editor(
				pd.DataFrame(base_rows),
				use_container_width=True,
				num_rows="dynamic",
				key=f"editor_indicadores_unidad_{unidad_id}",
			)

			if st.button("Guardar indicadores de unidad", use_container_width=True):
				filas = editor_indicadores.to_dict(orient="records")
				codigos = [str(f.get("codigo", "")).strip() for f in filas if str(f.get("codigo", "")).strip()]
				if not codigos:
					st.error("Debes definir al menos un indicador para la unidad.")
				elif len(codigos) != len(set(codigos)):
					st.error("Existen codigos de indicadores duplicados. Corrigelos antes de guardar.")
				else:
					reemplazar_indicadores_unidad(unidad_id, filas)
					st.success("Indicadores de unidad actualizados.")
					st.rerun()

	indicadores_db = listar_indicadores_unidad(unidad_id)
	codigos_indicadores = [row["codigo"] for row in indicadores_db]

	with st.container(border=True):
		st.markdown("### Fase 2: Carga de Instrumento")
		nombre_eval = st.text_input(
			"Nombre de evaluación",
			placeholder="Solemne 1",
			key=f"nombre_eval_{unidad_id}",
		)
		tab_archivo, tab_manual = st.tabs(["📄 Subir Evaluación (.docx / .pdf / .txt)", "✏️ Ingreso Manual de Preguntas"])

		with tab_archivo:
			archivo_eval = st.file_uploader(
				"Subir Evaluación (.docx / .pdf / .txt)",
				type=["docx", "pdf", "txt"],
				key=f"archivo_eval_{unidad_id}",
			)
			archivo_eval_state_key = f"archivo_eval_state_{unidad_id}"
			if archivo_eval is not None:
				st.session_state[archivo_eval_state_key] = {
					"name": archivo_eval.name,
					"bytes": archivo_eval.getvalue(),
				}

			archivo_eval_guardado = st.session_state.get(archivo_eval_state_key)
			preguntas_detectadas_archivo: list[dict[str, str]] = []
			if archivo_eval_guardado:
				with st.spinner("Leyendo y detectando preguntas del documento Word..."):
					try:
						nombre_archivo_guardado = str(archivo_eval_guardado.get("name", "")).lower()
						bytes_archivo_guardado = archivo_eval_guardado.get("bytes", b"")
						if nombre_archivo_guardado.endswith(".docx"):
							preguntas_detectadas_archivo = _extraer_preguntas_docx_estructuradas(bytes_archivo_guardado)
							texto_instrumento = _leer_texto_docx_desde_bytes(bytes_archivo_guardado)
						elif nombre_archivo_guardado.endswith(".pdf"):
							texto_instrumento = _leer_texto_pdf_desde_bytes(bytes_archivo_guardado)
							preguntas_detectadas_archivo = _detectar_preguntas_instrumento_detalladas(texto_instrumento)
						else:
							texto_instrumento = bytes_archivo_guardado.decode("utf-8", errors="ignore")
							preguntas_detectadas_archivo = _detectar_preguntas_instrumento_detalladas(texto_instrumento)
					except Exception as e:
						st.error(f"❌ Error al procesar el archivo: {e}")
						preguntas_detectadas_archivo = []

				with st.expander("👁️ Preguntas/Ítems detectados"):
					if preguntas_detectadas_archivo:
						st.write("\n".join(f"- {pregunta['display']}" for pregunta in preguntas_detectadas_archivo))
					else:
						st.info("No se detectaron preguntas/ítems en el archivo seleccionado.")

			if st.button("⚡ Procesar Evaluación e Ir a Matriz", type="primary", key="btn_procesar_archivo"):
				if not nombre_eval.strip():
					st.error("Debes indicar el nombre de la evaluacion.")
				elif archivo_eval is None and not archivo_eval_guardado:
					st.error("Debes subir un archivo de evaluación.")
				else:
					archivo_fuente = archivo_eval_guardado if archivo_eval_guardado else {"name": archivo_eval.name, "bytes": archivo_eval.getvalue()}
					nombre_archivo = str(archivo_fuente.get("name", "evaluacion.txt"))
					bytes_archivo = archivo_fuente.get("bytes", b"")
					with st.spinner("Leyendo y detectando preguntas del documento Word..."):
						try:
							if nombre_archivo.lower().endswith(".docx"):
								preguntas_detectadas_archivo = _extraer_preguntas_docx_estructuradas(bytes_archivo)
								texto_instrumento = _leer_texto_docx_desde_bytes(bytes_archivo)
							elif nombre_archivo.lower().endswith(".pdf"):
								texto_instrumento = _leer_texto_pdf_desde_bytes(bytes_archivo)
								preguntas_detectadas_archivo = _detectar_preguntas_instrumento_detalladas(texto_instrumento)
							else:
								texto_instrumento = bytes_archivo.decode("utf-8", errors="ignore")
								preguntas_detectadas_archivo = _detectar_preguntas_instrumento_detalladas(texto_instrumento)
						except Exception as e:
							st.error(f"❌ Error al procesar el archivo: {e}")
							preguntas_detectadas_archivo = []

					if not preguntas_detectadas_archivo:
						st.error("No se detectaron preguntas válidas en el archivo.")
						return

					eval_dir = _eval_unidad_dir(str(unidad_id))
					eval_dir.mkdir(parents=True, exist_ok=True)
					destino = eval_dir / nombre_archivo
					destino.write_bytes(bytes_archivo)
					evaluacion_id = guardar_evaluacion(unidad_id, nombre_eval.strip(), str(destino))
					st.session_state.evaluacion_activa_id = evaluacion_id
					st.session_state[archivo_eval_state_key] = {"name": nombre_archivo, "bytes": bytes_archivo}
					st.success("Evaluación registrada y lista para construir la matriz.")
					st.rerun()

		with tab_manual:
			texto_preguntas_eval = st.text_area(
				"📝 Ingrese las preguntas de la evaluación (Una pregunta por línea)",
				height=140,
				placeholder="Ingrese una pregunta por línea o separadas por un entero...",
				key="texto_preguntas_eval",
			)
			preguntas_detectadas_manual = _detectar_preguntas_instrumento_detalladas(texto_preguntas_eval, permitir_secuencial=True)

			with st.expander("👁️ Preguntas/Ítems detectados"):
				if preguntas_detectadas_manual:
					st.write("\n".join(f"- {pregunta['display']}" for pregunta in preguntas_detectadas_manual))
				else:
					st.info("Ingresa texto para detectar preguntas automáticamente.")

			if st.button("📝 Registrar Preguntas Manuales", key="btn_procesar_manual"):
				if not nombre_eval.strip():
					st.error("Debes indicar el nombre de la evaluacion.")
				elif not texto_preguntas_eval.strip():
					st.error("Debes ingresar al menos una pregunta en la evaluación.")
				elif not preguntas_detectadas_manual:
					st.error("No se detectaron preguntas válidas en el texto manual.")
				else:
					eval_dir = _eval_unidad_dir(str(unidad_id))
					eval_dir.mkdir(parents=True, exist_ok=True)
					slug_eval = _slug(nombre_eval.strip()) or "evaluacion"
					destino_txt = eval_dir / f"{slug_eval}.txt"
					destino_txt.write_text(texto_preguntas_eval.strip(), encoding="utf-8")
					evaluacion_id = guardar_evaluacion(unidad_id, nombre_eval.strip(), str(destino_txt))
					st.session_state.evaluacion_activa_id = evaluacion_id
					st.success("Preguntas manuales registradas y listas para la matriz.")
					st.rerun()

	evaluaciones = listar_evaluaciones_unidad(unidad_id)
	if not evaluaciones:
		st.info("Aun no hay evaluaciones registradas para esta unidad.")
		return

	with st.container(border=True):
		st.markdown("### Fase 3: Mapeo Asistido por IA")
		mapa_eval = {f"{e['nombre']} (ID {e['id']})": e for e in evaluaciones}
		opciones_eval = list(mapa_eval.keys())
		idx_eval = 0
		for i, key in enumerate(opciones_eval):
			if int(mapa_eval[key]["id"]) == int(st.session_state.evaluacion_activa_id or -1):
				idx_eval = i
				break
		seleccion_eval = st.selectbox("Evaluacion activa", opciones_eval, index=idx_eval)
		evaluacion = mapa_eval[seleccion_eval]
		evaluacion_id = int(evaluacion["id"])
		st.session_state.evaluacion_activa_id = evaluacion_id

		confirmar_eliminacion_eval = st.checkbox(
			"Confirmar eliminacion de evaluacion activa",
			key=f"confirm_del_eval_{evaluacion_id}",
		)
		if st.button("🗑️ Eliminar evaluación activa", use_container_width=True, type="secondary"):
			if not confirmar_eliminacion_eval:
				st.warning("Marca la confirmacion antes de eliminar la evaluacion activa.")
			else:
				eliminar_evaluacion(evaluacion_id)
				if st.session_state.evaluacion_activa_id == evaluacion_id:
					st.session_state.evaluacion_activa_id = None
				st.success("Evaluacion eliminada correctamente.")
				st.rerun()

		archivo_path = Path(str(evaluacion.get("archivo_path", "")).strip())
		mapeo_existente = listar_mapeo_evaluacion(evaluacion_id)

		propuesta_rows: list[dict[str, Any]] = []
		if mapeo_existente:
			propuesta_rows = _consolidar_pregunta_indicadores([
				{
					"numero_pregunta": row["numero_pregunta"],
					"puntos_maximos": row["puntos_maximos"],
					"indicadores_asociados": row.get("indicadores_asociados", ""),
				}
				for row in mapeo_existente
			])
		elif archivo_path.exists():
			try:
				texto_instrumento = _leer_texto_instrumento_desde_path(archivo_path)
				permitir_secuencial = archivo_path.suffix.lower() == ".txt"
				if archivo_path.suffix.lower() == ".docx":
					lineas_preguntas_detalladas = _extraer_preguntas_docx_estructuradas(archivo_path.read_bytes())
				else:
					lineas_preguntas_detalladas = _detectar_preguntas_instrumento_detalladas(texto_instrumento, permitir_secuencial=permitir_secuencial)
			except Exception as exc:
				st.error(f"No se pudo leer el archivo de preguntas: {exc}")
				lineas_preguntas_detalladas = []

			if not lineas_preguntas_detalladas:
				st.warning("El archivo de preguntas no contiene lineas validas para construir la matriz base.")
			else:
				autoseed_indicador = codigos_indicadores[0] if codigos_indicadores else ""
				propuesta_rows = [
					{
						"numero_pregunta": fila["codigo"],
						"puntos_maximos": 1.0,
						"indicadores_asociados": autoseed_indicador,
					}
					for fila in lineas_preguntas_detalladas
				]
		else:
			st.error("No se encontro el archivo txt de preguntas para construir la matriz base.")
			propuesta_rows = []

		editor_mapeo = st.data_editor(
			pd.DataFrame(propuesta_rows),
			use_container_width=True,
			num_rows="fixed",
			key=f"editor_mapeo_eval_{evaluacion_id}",
			column_config={
				"numero_pregunta": st.column_config.TextColumn(disabled=True),
				"puntos_maximos": st.column_config.NumberColumn(min_value=0.1, step=0.1),
				"indicadores_asociados": st.column_config.TextColumn(
					help="Ingrese los códigos separados por coma (ej. 1.1, 4.3.3)"
				),
			},
		)

		indicadores_usados: set[str] = set()
		for fila in editor_mapeo.to_dict(orient="records"):
			for codigo in _parsear_indicadores_asociados(str(fila.get("indicadores_asociados", ""))):
				indicadores_usados.add(codigo)
		indicadores_faltantes = set(codigos_indicadores) - indicadores_usados

		if indicadores_faltantes:
			st.warning(
				"⚠️ Alerta de Cobertura Doctrinal: Faltan preguntas para evaluar los siguientes indicadores: "
				+ ", ".join(sorted(indicadores_faltantes))
			)
		else:
			st.success("✅ Cobertura del 100%. Todos los indicadores están siendo evaluados.")

		if st.button("Confirmar y Validar Matriz de Evaluación", use_container_width=True, type="primary"):
			filas = editor_mapeo.to_dict(orient="records")
			filas_validas_raw: list[dict[str, Any]] = []
			for fila in filas:
				numero = str(fila.get("numero_pregunta", "")).strip()
				indicadores_texto = _normalizar_lista_indicadores(
					str(fila.get("indicadores_asociados", ""))
				)
				try:
					puntos = float(fila.get("puntos_maximos", 0))
				except (TypeError, ValueError):
					puntos = 0.0

				if not numero or not indicadores_texto or puntos <= 0:
					continue
				if codigos_indicadores:
					indices = _parsear_indicadores_asociados(indicadores_texto)
					if not indices or any(ind not in codigos_indicadores for ind in indices):
						continue

				filas_validas_raw.append(
					{
						"numero_pregunta": numero.upper(),
						"puntos_maximos": puntos,
						"indicadores_asociados": indicadores_texto,
					}
				)

			filas_validas = _consolidar_pregunta_indicadores(filas_validas_raw)

			if not filas_validas:
				st.error("La matriz no tiene filas validas para guardar.")
			else:
				reemplazar_mapeo_evaluacion(evaluacion_id, filas_validas)
				st.success("Matriz de evaluacion validada y guardada.")
				st.rerun()


def _reindexar_unidad(unidad_id: str) -> None:
	doc_dir = _data_unidad_dir(unidad_id)
	vector_dir = _vector_unidad_dir(unidad_id)
	doc_dir.mkdir(parents=True, exist_ok=True)

	docs = sorted(
		p
		for p in doc_dir.iterdir()
		if p.is_file() and p.suffix.lower() in {".docx", ".pdf", ".pptx", ".ppt"}
	)
	if not docs:
		shutil.rmtree(vector_dir, ignore_errors=True)
		return

	documentos = cargar_documentos_docx(origen=doc_dir)
	construir_vector_store(
		documentos=documentos,
		persist_directory=vector_dir,
		collection_name=f"acapomil_{unidad_id}",
	)


def _vista_bibliografia() -> None:
	st.subheader("📖 Bibliografia dinamica por Unidad Activa")
	with st.container(border=True):
		st.info("Instrucciones de uso: carga archivos DOCX, PDF o PowerPoint, vectoriza y administra el repositorio documental de la unidad activa.")
	unidad = _unidad_activa()
	if not unidad:
		st.info("Define y activa una unidad para gestionar doctrina.")
		return

	unidad_id = str(unidad["id"])
	doc_dir = _data_unidad_dir(unidad_id)
	doc_dir.mkdir(parents=True, exist_ok=True)

	with st.container(border=True):
		st.markdown("### ⬆️ Cargar documentos doctrinales")
		archivos = st.file_uploader(
			"Cargar doctrina",
			type=["docx", "pdf", "pptx", "ppt"],
			accept_multiple_files=True,
			key=f"upload_{unidad_id}",
		)
		if st.button("Guardar y vectorizar", use_container_width=True, type="primary"):
			if not archivos:
				st.warning("Selecciona al menos un documento DOCX, PDF, PPTX o PPT")
			else:
				for archivo in archivos:
					destino = doc_dir / archivo.name
					destino.write_bytes(archivo.getvalue())
				with st.spinner("Vectorizando bibliografia de la unidad activa..."):
					try:
						_reindexar_unidad(unidad_id)
						st.success("Documentos guardados y vector store actualizado.")
					except Exception as exc:
						st.error(f"No se pudo vectorizar: {exc}")

	docs = sorted(
		p
		for p in doc_dir.iterdir()
		if p.is_file() and p.suffix.lower() in {".docx", ".pdf", ".pptx", ".ppt"}
	)
	with st.container(border=True):
		st.markdown("### 🗂️ Documentos actuales")
		if not docs:
			st.info("No hay documentos cargados para esta unidad.")
		else:
			tabla = pd.DataFrame(
				[
					{
						"Documento": p.name,
						"Tamano (KB)": round(p.stat().st_size / 1024, 2),
					}
					for p in docs
				]
			)
			st.dataframe(tabla, use_container_width=True)

			archivo_del = st.selectbox("Documento a eliminar", [p.name for p in docs])
			c1, c2 = st.columns(2)
			with c1:
				if st.button("Eliminar documento", use_container_width=True):
					(doc_dir / archivo_del).unlink(missing_ok=True)
					try:
						_reindexar_unidad(unidad_id)
					except Exception as exc:
						st.error(f"Documento eliminado, pero fallo reindexacion: {exc}")
					st.success("Documento eliminado.")
					st.rerun()
			with c2:
				if st.button("Eliminar vector store de la unidad", use_container_width=True):
					shutil.rmtree(_vector_unidad_dir(unidad_id), ignore_errors=True)
					st.warning("Vector store eliminado para la unidad activa.")


def _alumnos_unidad(unidad_id: str) -> list[dict[str, Any]]:
	unidad = obtener_unidad_por_id(int(unidad_id)) if str(unidad_id).isdigit() else None
	if not unidad:
		return []
	# Xuanwu: consumo docente centralizado por filtro exacto CRIM + Mencion.
	return listar_alumnos(anio=unidad["anio"], mencion=unidad["mencion"])


def _vista_alumnos() -> None:
	st.subheader("🧑‍🎓 Alumnos de la Unidad Activa")
	with st.container(border=True):
		st.info("Instrucciones de uso: visualiza los alumnos matriculados por Administracion para el CRIM y Mencion de la unidad activa.")
	unidad = _unidad_activa()
	if not unidad:
		st.info("Activa una unidad para visualizar alumnos filtrados automaticamente.")
		return
	alumnos = _alumnos_unidad(str(unidad["id"]))
	with st.container(border=True):
		st.markdown("### 📋 Alumnos de la unidad activa")
		if not alumnos:
			st.info("Aun no hay alumnos registrados en esta unidad.")
			return

		# Xuanwu: mostramos solo columnas pedagogicas, ocultando ids tecnicos.
		tabla_alumnos = pd.DataFrame(alumnos)
		if not tabla_alumnos.empty:
			tabla_alumnos = tabla_alumnos[["rut", "nombre", "anio", "mencion"]]
		st.dataframe(tabla_alumnos, use_container_width=True)
		st.caption("Gestion de altas/bajas restringida al Administrador.")


def _default_result_rows() -> list[dict[str, Any]]:
	return [
		{"Indicador": "R4.1.1", "Logro": 0.0, "Observacion": ""},
		{"Indicador": "R4.2.1", "Logro": 0.0, "Observacion": ""},
	]


def _vista_trazabilidad() -> None:
	st.subheader("🧠 Flujo de Trazabilidad RAG")
	with st.container(border=True):
		st.info("Instrucciones de uso: ingresa puntajes por pregunta de la evaluacion activa, calcula logro por indicador y prioriza brechas para el plan RAG.")
	unidad = _unidad_activa()
	if not unidad:
		st.info("Activa una unidad para habilitar trazabilidad.")
		return

	evaluaciones = listar_evaluaciones_unidad(int(unidad["id"]))
	if not evaluaciones:
		st.warning("No existen evaluaciones diseñadas. Usa primero la pestaña Diseño de Evaluación.")
		return

	mapa_eval = {f"{e['nombre']} (ID {e['id']})": e for e in evaluaciones}
	opciones_eval = list(mapa_eval.keys())
	idx_eval = 0
	for i, key in enumerate(opciones_eval):
		if int(mapa_eval[key]["id"]) == int(st.session_state.evaluacion_activa_id or -1):
			idx_eval = i
			break
	seleccion_eval = st.selectbox("Evaluacion para trazabilidad", opciones_eval, index=idx_eval)
	evaluacion = mapa_eval[seleccion_eval]
	evaluacion_id = int(evaluacion["id"])
	st.session_state.evaluacion_activa_id = evaluacion_id

	mapeo_eval = listar_mapeo_evaluacion(evaluacion_id)
	if not mapeo_eval:
		st.warning("La evaluacion seleccionada no tiene matriz validada. Completa la Fase 3 en Diseño de Evaluación.")
		return
	mapeo_eval = _consolidar_pregunta_indicadores([
		{
			"numero_pregunta": row["numero_pregunta"],
			"puntos_maximos": row["puntos_maximos"],
			"indicadores_asociados": row.get("indicadores_asociados", ""),
		}
		for row in mapeo_eval
	])

	alumnos = listar_alumnos(anio=unidad["anio"], mencion=unidad["mencion"])
	if not alumnos:
		st.warning("Registra alumnos primero para continuar.")
		return

	st.markdown("### 1) Ingreso de puntajes por pregunta")
	alumno_label_map = {f"{a['nombre']} ({a['rut']})": a for a in alumnos}
	alumno_sel_label = st.selectbox("Alumno", list(alumno_label_map.keys()))
	alumno_sel = alumno_label_map[alumno_sel_label]
	alumno_id = int(alumno_sel["id"])
	puntajes_historicos = obtener_puntajes_pregunta(alumno_id=alumno_id, evaluacion_id=evaluacion_id)

	df_preguntas = pd.DataFrame(
		{
			"numero_pregunta": row["numero_pregunta"],
			"indicadores_asociados": row.get("indicadores_asociados", ""),
			"puntos_maximos": float(row["puntos_maximos"]),
			"puntos_obtenidos": float(puntajes_historicos.get(str(row["numero_pregunta"]), 0.0)),
		}
		for row in mapeo_eval
	)
	if not df_preguntas.empty:
		df_preguntas = df_preguntas.sort_values(
			by="numero_pregunta",
			key=lambda x: x.map(_clave_orden_pregunta),
		).reset_index(drop=True)

	edited = st.data_editor(
		df_preguntas,
		use_container_width=True,
		num_rows="fixed",
		key=f"editor_matriz_traza_{alumno_id}_{evaluacion_id}",
		column_config={
			"puntos_obtenidos": st.column_config.NumberColumn(min_value=0.0, step=0.1),
			"puntos_maximos": st.column_config.NumberColumn(disabled=True),
			"indicadores_asociados": st.column_config.TextColumn(disabled=True),
			"numero_pregunta": st.column_config.TextColumn(disabled=True),
		},
	)

	resumen_indicadores: list[dict[str, Any]] = []
	acumulado: dict[str, dict[str, float]] = {}
	for fila in edited.to_dict(orient="records"):
		indicadores = _parsear_indicadores_asociados(str(fila.get("indicadores_asociados", "")))
		try:
			maximo = float(fila.get("puntos_maximos", 0.0))
		except (TypeError, ValueError):
			maximo = 0.0
		try:
			obtenido = float(fila.get("puntos_obtenidos", 0.0))
		except (TypeError, ValueError):
			obtenido = 0.0

		if maximo < 0:
			maximo = 0.0
		obtenido = max(0.0, min(obtenido, maximo if maximo > 0 else obtenido))

		for indicador in indicadores:
			if indicador not in acumulado:
				acumulado[indicador] = {"obtenido": 0.0, "maximo": 0.0}
			acumulado[indicador]["obtenido"] += obtenido
			acumulado[indicador]["maximo"] += maximo

	for indicador, valores in acumulado.items():
		maximo = float(valores["maximo"])
		obtenido = float(valores["obtenido"])
		logro = (obtenido / maximo) if maximo > 0 else 0.0
		resumen_indicadores.append(
			{
				"Indicador": indicador,
				"Logro": round(logro, 4),
				"Puntaje": f"{obtenido:.2f}/{maximo:.2f}",
			}
		)

	if st.button("Guardar resultados por matriz", use_container_width=True):
		guardar_puntajes_pregunta(
			alumno_id=alumno_id,
			evaluacion_id=evaluacion_id,
			puntajes=edited.to_dict(orient="records"),
		)
		registros = [
			{
				"Indicador": fila["Indicador"],
				"Logro": fila["Logro"],
				"Observacion": f"{evaluacion['nombre']} | {fila['Puntaje']}",
			}
			for fila in resumen_indicadores
		]
		reemplazar_resultados_alumno(alumno_id=alumno_id, resultados=registros)
		st.success("Resultados por indicador guardados desde la matriz de evaluacion.")
		st.rerun()

	if resumen_indicadores:
		st.markdown("### 2) Logro consolidado por indicador")
		st.dataframe(pd.DataFrame(resumen_indicadores), use_container_width=True)

	mapa_cohortes: dict[str, list[str]] = {}
	for alumno in alumnos:
		resultados = listar_resultados_alumno(int(alumno["id"]))
		if not resultados:
			continue
		nombre_alumno = str(alumno["nombre"]).strip()
		for fila in resultados:
			try:
				if float(fila.get("logro", 0)) >= 0.60:
					continue
			except (TypeError, ValueError):
				continue
			indicador = str(fila.get("indicador", "")).strip()
			if not indicador:
				continue
			claves = mapa_cohortes.setdefault(indicador, [])
			if nombre_alumno and nombre_alumno not in claves:
				claves.append(nombre_alumno)

	with st.container(border=True):
		st.markdown("### 📊 Diagnosticó Colectivo y Agrupación por Cohorte")
		if not mapa_cohortes:
			st.info("No se detectaron cohortes con indicadores bajo 60% en la unidad activa.")
		else:
			for indicador, nombres in sorted(mapa_cohortes.items(), key=lambda item: item[0]):
				st.markdown(f"**{indicador}**")
				st.write(", ".join(nombres))

			opciones_cohortes = [
				f"{indicador} ({len(nombres)} cadetes)"
				for indicador, nombres in sorted(mapa_cohortes.items(), key=lambda item: item[0])
			]
			indice_cohorte = 0
			if opciones_cohortes:
				seleccion_cohorte = st.selectbox(
					"Selector de grupo/cohorte",
					opciones_cohortes,
					index=indice_cohorte,
					key=f"selector_cohorte_{evaluacion_id}",
				)
				indicador_seleccionado = seleccion_cohorte.split(" (")[0]
				cadetes_cohorte = mapa_cohortes.get(indicador_seleccionado, [])
				tau_cohorte = st.slider(
					"Tau retrieval para cohorte",
					0.50,
					0.95,
					0.75,
					0.01,
					key=f"tau_cohorte_{evaluacion_id}",
				)

				if st.button(
					"Generar Plan Colectivo de 4 Semanas para la Cohorte Seleccionada",
					use_container_width=True,
					type="primary",
				):
					if not cadetes_cohorte:
						st.warning("La cohorte seleccionada no tiene cadetes asociados.")
					else:
						with st.spinner("Generando plan colectivo para la cohorte seleccionada..."):
							try:
								salida_cohorte = generar_plan_trabajo_cohorte(
									cadetes_cohorte,
									brechas_forzadas=[indicador_seleccionado],
									tau=tau_cohorte,
								)
								st.session_state.borrador_retro = str(salida_cohorte.get("plan_4_semanas", ""))
								st.success("Plan colectivo generado y cargado en el editor de borrador.")
							except Exception as exc:
								st.error(f"No se pudo generar el plan colectivo: {exc}")

	brechas_locales = []
	for fila in resumen_indicadores:
		try:
			if float(fila.get("Logro", 0)) < 0.60:
				brechas_locales.append(str(fila.get("Indicador", "")))
		except (TypeError, ValueError):
			continue

	if brechas_locales:
		st.warning(f"Brechas detectadas (<0.60): {', '.join(brechas_locales)}")
	else:
		st.info("Sin brechas locales detectadas para el alumno seleccionado.")

	st.markdown("### 3) Generar retroalimentacion asistida")
	col1, col2 = st.columns([2, 1])
	with col1:
		alumno_brecha = alumno_sel["nombre"]
		st.text_input("Alumno objetivo (fuente trazabilidad local)", value=alumno_brecha, disabled=True)
	with col2:
		tau = st.slider("Tau retrieval", 0.50, 0.95, 0.75, 0.01)

	if st.button("Generar Plan de 4 Semanas", use_container_width=True, type="primary"):
		if not brechas_locales:
			st.warning("No hay brechas bajo 60% en la matriz actual para este alumno.")
			return
		with st.spinner("Invocando orquestador RAG + LLM..."):
			try:
				# Xuanwu: integramos brechas locales de la matriz con el orquestador oficial.
				salida = generar_plan_trabajo_alumno(
					alumno_brecha,
					tau=tau,
					brechas_forzadas=brechas_locales,
				)
				st.session_state.borrador_retro = str(salida.get("plan_4_semanas", ""))
				st.success("Borrador generado. Puedes editarlo antes de emitirlo.")
			except Exception as exc:
				st.error(f"No se pudo generar el plan: {exc}")

	tab_editar, tab_vista_previa = st.tabs([
		"📝 Editar Plan (Texto Plano)",
		"👁️ Vista Previa de Alto Estándar",
	])

	with tab_editar:
		st.text_area(
			"Borrador editable (Human-in-the-loop)",
			key="borrador_retro",
			height=400,
			placeholder="El plan de 4 semanas aparecerá aquí para revisión docente.",
		)

	with tab_vista_previa:
		contenido_actual = st.session_state.borrador_retro
		if not contenido_actual.strip():
			st.info("Aún no se ha generado ningún plan. Presiona el botón de generación para empezar.")
		else:
			st.markdown("### 👁️ Previsualización del Plan de Trabajo")
			st.markdown(contenido_actual, unsafe_allow_html=True)

	if st.button("💾 Guardar Plan en Historial", use_container_width=True):
		if not st.session_state.borrador_retro.strip():
			st.warning("No hay contenido para guardar en el historial.")
		else:
			ok = guardar_retroalimentacion_local(
				rut_alumno=str(alumno_sel["rut"]),
				unidad_id=str(unidad["id"]),
				plan_texto=st.session_state.borrador_retro,
				tau=tau,
			)
			if ok:
				st.success("Plan guardado en el historial local.")
				st.rerun()
			else:
				st.error("No se pudo guardar el plan en el historial local.")

	with st.expander("📂 Ver Historial de Planes Guardados", expanded=False):
		historial = obtener_historial_retroalimentaciones(
			rut_alumno=str(alumno_sel["rut"]),
			unidad_id=str(unidad["id"]),
		)
		if not historial:
			st.info("No hay planes guardados para este alumno en la unidad activa.")
		else:
			mapa_historial = {
				f"{item['fecha_generacion']} | tau={item['tau_utilizado']}": item
				for item in historial
			}

			opciones_historial = list(mapa_historial.keys())

			# Declaración segura del selectbox sin asignaciones automáticas directas
			seleccion_historial = st.selectbox(
				"Selecciona un plan histórico",
				opciones_historial,
				key=f"historial_retro_{alumno_id}_{unidad['id']}",
			)

			# Control defensivo estricto antes de renderizar los botones de acción
			if seleccion_historial and seleccion_historial in mapa_historial:
				plan_seleccionado = mapa_historial[seleccion_historial]
				col_cargar, col_eliminar = st.columns(2)

				with col_cargar:
					if st.button("🔌 Cargar plan histórico", use_container_width=True):
						# La asignación del texto plano sin formato ocurre SOLAMENTE en el click del botón
						st.session_state.borrador_retro = str(plan_seleccionado["contenido_plan"])
						st.success("Plan histórico cargado en el editor.")
						st.rerun()

				with col_eliminar:
					if st.button("🗑️ Eliminar de historial", use_container_width=True, type="secondary"):
						retro_id = int(plan_seleccionado["id"])
						if eliminar_retroalimentacion_local(retro_id):
							st.success("✅ Plan histórico eliminado del registro.")
							st.rerun()
						else:
							st.error("No se pudo eliminar el plan del historial.")
			else:
				st.caption("Actualizando historial para el estudiante seleccionado...")


def _vista_reportes() -> None:
	st.subheader("📊 Reporte de Jefatura")
	with st.container(border=True):
		st.info("Instrucciones de uso: revisa metricas de riesgo bajo 60% y utiliza las alertas para priorizar acompanamiento docente.")
	unidad = _unidad_activa()
	if not unidad:
		st.info("Activa una unidad para generar reportes.")
		return

	alumnos = listar_alumnos(anio=unidad["anio"], mencion=unidad["mencion"])
	total = len(alumnos)

	bajo_60_local = 0
	detalle_local: list[dict[str, Any]] = []
	conteo_brechas_indicador: dict[str, int] = {}
	logros_bajo_umbral: list[float] = []
	for alumno in alumnos:
		filas = listar_resultados_alumno(int(alumno["id"]))
		brechas = []
		for fila in filas:
			try:
				if float(fila.get("logro", 0)) < 0.60:
					indicador = str(fila.get("indicador", ""))
					brechas.append(indicador)
					conteo_brechas_indicador[indicador] = conteo_brechas_indicador.get(indicador, 0) + 1
					logros_bajo_umbral.append(float(fila.get("logro", 0)))
			except (TypeError, ValueError):
				continue
		if brechas:
			bajo_60_local += 1
			detalle_local.append(
				{
					"Alumno": alumno["nombre"],
					"RUT": alumno["rut"],
					"Brechas": ", ".join(brechas),
				}
			)

	bajo_60_excel = 0
	try:
		for nombre in obtener_alumnos():
			if obtener_brechas_alumno(nombre).get("indicadores_brecha"):
				bajo_60_excel += 1
	except Exception:
		pass

	porcentaje_riesgo_local = (bajo_60_local / total * 100.0) if total else 0.0
	logro_promedio_bajo = (sum(logros_bajo_umbral) / len(logros_bajo_umbral) * 100.0) if logros_bajo_umbral else 0.0
	delta_logro = f"-{100.0 - logro_promedio_bajo:.1f}% de logro promedio" if logros_bajo_umbral else "Sin rezago critico"

	if bajo_60_local > 0:
		st.error(
			f"⚠️ Alerta: Se ha detectado {bajo_60_local} estudiante"
			f"{'s' if bajo_60_local != 1 else ''} con brechas por debajo del umbral mínimo del 60%. "
			"Requiere acompañamiento pedagógico."
		)

	c1, c2, c3 = st.columns(3)
	with c1:
		with st.container(border=True):
			st.metric("Alumnos en unidad activa", total)
	with c2:
		with st.container(border=True):
			st.metric(
				"Riesgo Local",
				bajo_60_local,
				delta=delta_logro,
			)
			if bajo_60_local > 0:
				st.warning(f"{porcentaje_riesgo_local:.1f}% del curso presenta alertas locales.")
	with c3:
		with st.container(border=True):
			st.metric("Riesgo Excel", bajo_60_excel)

	with st.container(border=True):
		st.markdown("### 📊 Diagnóstico Colectivo de Indicadores")
		if not conteo_brechas_indicador:
			st.info("No se detectan brechas colectivas bajo 60% en la unidad activa.")
		else:
			df_brechas = pd.DataFrame(
				[
					{"Indicador": indicador, "Alumnos con brecha": cantidad}
					for indicador, cantidad in conteo_brechas_indicador.items()
				]
			)
			df_brechas = df_brechas.sort_values(by="Alumnos con brecha", ascending=False)
			try:
				import plotly.express as px

				fig = px.bar(
					df_brechas,
					x="Alumnos con brecha",
					y="Indicador",
					orientation="h",
					text="Alumnos con brecha",
				)
				fig.update_layout(height=360, margin=dict(l=20, r=20, t=20, b=20))
				st.plotly_chart(fig, use_container_width=True)
			except Exception:
				st.bar_chart(df_brechas.set_index("Indicador"), use_container_width=True)

	with st.container(border=True):
		with st.expander("🔍 Detalle Nominal de Alumnos en Riesgo", expanded=False):
			if not detalle_local:
				st.success("Sin alertas locales para la unidad activa.")
			else:
				st.dataframe(pd.DataFrame(detalle_local), use_container_width=True)


def main() -> None:
	_inject_styles()
	_init_state()

	if not st.session_state.is_logged:
		_render_login()
		return

	_render_top_context()
	modulo = _menu_sidebar()

	if st.session_state.usuario_rol == "admin":
		if modulo.startswith("1"):
			_vista_admin_matricula()
		else:
			_vista_admin_cuentas()
		return

	if modulo.startswith("1"):
		_vista_unidades()
	elif modulo.startswith("2"):
		_vista_diseno_evaluacion()
	elif modulo.startswith("3"):
		_vista_bibliografia()
	elif modulo.startswith("4"):
		_vista_trazabilidad()
	elif modulo.startswith("5"):
		_vista_reportes()


if __name__ == "__main__":
	main()
