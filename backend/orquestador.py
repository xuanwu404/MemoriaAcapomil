from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import os

import numpy as np
if not hasattr(np, "float_"):
	np.float_ = np.float64

from dotenv import load_dotenv
from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.runnables import RunnableLambda

from backend.database import obtener_brechas_alumno
from backend.pipeline_rag import (
	TAU_SIMILITUD,
	ingest_arreglo_doctrina,
	recuperar_fragmentos_relevantes,
)


load_dotenv()


DEFAULT_OLLAMA_MODEL = "llama3.1"
DEFAULT_TAU = TAU_SIMILITUD


@dataclass(frozen=True)
class ContextoRefuerzo:
	indicador: str
	fragmentos: list[str]


def _crear_llm():
	"""Selecciona LLM segun LLM_PROVIDER del .env."""
	provider = os.getenv("LLM_PROVIDER", "ollama").strip().lower()

	if provider == "ollama":
		try:
			from langchain_community.chat_models import ChatOllama
		except ImportError as exc:
			raise ImportError(
				"No se encontro ChatOllama. Instala langchain-community actualizado "
				"o langchain-ollama para usar LLM_PROVIDER=ollama."
			) from exc

		return ChatOllama(
			model=DEFAULT_OLLAMA_MODEL,
			base_url="http://localhost:11434",
			temperature=0.2,
		)

	if provider == "openai":
		try:
			from langchain_openai import ChatOpenAI
		except ImportError as exc:
			raise ImportError(
				"No se encontro ChatOpenAI. Instala langchain-openai para usar "
				"LLM_PROVIDER=openai."
			) from exc

		return ChatOpenAI(
			model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
			temperature=0.2,
		)

	raise ValueError(
		"LLM_PROVIDER invalido. Valores soportados: 'ollama' u 'openai'."
	)


def _asegurar_retrieval_disponible() -> None:
	"""Crea el vector store si aun no existe para no romper la cadena."""
	try:
		# Xuanwu: esta llamada falla si no hay vector store persistido.
		recuperar_fragmentos_relevantes("test de disponibilidad", k=1)
	except FileNotFoundError:
		# Xuanwu: auto-ingesta para mantener experiencia MVP sin pasos manuales.
		ingest_arreglo_doctrina()


def _recuperar_contexto_por_indicador(indicador: str, tau: float, unidad_id: str = None) -> ContextoRefuerzo:
	consulta = f"Refuerzo doctrinario para mejorar el indicador {indicador} en evaluacion por competencias."

	if unidad_id:
		from pathlib import Path
		from backend.pipeline_rag import cargar_vector_store, recuperar_fragmentos_relevantes
		base_dir = Path(__file__).resolve().parent.parent / "data" / "vector_store" / str(unidad_id)
		col_name = f"acapomil_{unidad_id}"
		try:
			vstore = cargar_vector_store(persist_directory=base_dir, collection_name=col_name)
			resultados = recuperar_fragmentos_relevantes(pregunta=consulta, tau=tau, k=4, vector_store=vstore)
		except FileNotFoundError:
			resultados = []
	else:
		from backend.pipeline_rag import recuperar_fragmentos_relevantes
		resultados = recuperar_fragmentos_relevantes(pregunta=consulta, tau=tau, k=4)

	fragmentos = [item.fragmento for item in resultados]
	return ContextoRefuerzo(indicador=indicador, fragmentos=fragmentos)


def _formatear_contexto_refuerzo(contextos: list[ContextoRefuerzo]) -> str:
	bloques: list[str] = []
	for contexto in contextos:
		if contexto.fragmentos:
			for fragmento in contexto.fragmentos:
				bloques.append(f"- {fragmento}")

	if not bloques:
		return "No hay fragmentos específicos en la base vectorial. Basa las actividades estrictamente en tu conocimiento experto técnico sobre los temas indicados."

	return "\n".join(bloques)


def _normalizar_brechas(brechas_forzadas: list[str] | None) -> list[str]:
	return [str(item).strip() for item in brechas_forzadas or [] if str(item).strip()]


def _armar_entrada_llm(payload: dict[str, Any]) -> dict[str, Any]:
	from backend.database import obtener_brechas_alumno, listar_indicadores_unidad

	nombre_alumno = str(payload.get("nombre_alumno", "")).strip()
	alumnos_grupo = [str(item).strip() for item in payload.get("alumnos_grupo", []) if str(item).strip()]
	tau = float(payload.get("tau", DEFAULT_TAU))
	brechas_forzadas = payload.get("brechas_forzadas")
	modo_cohorte = bool(payload.get("modo_cohorte", False))
	unidad_id = payload.get("unidad_id")

	if modo_cohorte:
		if not alumnos_grupo:
			raise ValueError("Debe indicar una lista valida de alumnos_grupo.")
	else:
		if not nombre_alumno:
			raise ValueError("Debe indicar un nombre_alumno valido.")

	if brechas_forzadas is not None:
		indicadores = _normalizar_brechas(brechas_forzadas)
		resumen = {
			"alumno": nombre_alumno,
			"alumnos_grupo": alumnos_grupo,
			"puntaje_total": None,
			"nota_final": None,
			"indicadores_brecha": indicadores,
		}
	else:
		if modo_cohorte:
			brechas_unicas: list[str] = []
			for alumno in alumnos_grupo:
				for indicador in obtener_brechas_alumno(alumno).get("indicadores_brecha", []):
					indicador_limpio = str(indicador).strip()
					if indicador_limpio and indicador_limpio not in brechas_unicas:
						brechas_unicas.append(indicador_limpio)
			resumen = {
				"alumno": ", ".join(alumnos_grupo),
				"alumnos_grupo": alumnos_grupo,
				"puntaje_total": None,
				"nota_final": None,
				"indicadores_brecha": brechas_unicas,
			}
			indicadores = brechas_unicas
		else:
			resumen = obtener_brechas_alumno(nombre_alumno)
			indicadores = resumen.get("indicadores_brecha", [])

	# EXTRAER DESCRIPCIONES REALES DESDE LA BASE DE DATOS
	descripciones_temas = []
	if unidad_id:
		try:
			inds_db = listar_indicadores_unidad(int(unidad_id))
			mapa_inds = {r["codigo"]: r.get("descripcion", "") for r in inds_db}
			for cod in indicadores:
				desc = mapa_inds.get(cod, "")
				descripciones_temas.append(desc if desc else cod)
		except Exception:
			descripciones_temas = indicadores
	else:
		descripciones_temas = indicadores

	descripciones_temas = [t for t in descripciones_temas if t]

	contextos = [
		_recuperar_contexto_por_indicador(indicador=cod, tau=tau, unidad_id=unidad_id)
		for cod in indicadores
	]

	return {
		"alumno": resumen.get("alumno", nombre_alumno),
		"alumnos_grupo": resumen.get("alumnos_grupo", alumnos_grupo),
		"puntaje_total": resumen.get("puntaje_total"),
		"nota_final": resumen.get("nota_final"),
		"indicadores_brecha": ", ".join(descripciones_temas) if descripciones_temas else "Sin brechas",
		"contexto_refuerzo": _formatear_contexto_refuerzo(contextos),
		"tau": tau,
	}


def construir_cadena_plan_4_semanas(colectivo: bool = False):
	"""Construye una cadena LCEL para generar plan pedagogico en formato de tabla vertical, forzando contexto tecnico."""
	_asegurar_retrieval_disponible()

	instrucciones_base = (
		"Eres un docente militar del área de Ingeniería y Tecnología elaborando un Plan de Trabajo práctico y detallado.\n"
		"Datos:\n"
		"- Alumno(s): {alumno}\n"
		"- Temas técnicos a estudiar (Brechas reales detectadas): {indicadores_brecha}\n"
		"- Doctrina técnica de apoyo:\n{contexto_refuerzo}\n\n"
		"REGLAS ESTRICTAS Y OBLIGATORIAS:\n"
		"1) PROHIBIDO INVENTAR ESCENARIOS BÉLICOS: Basa las actividades ÚNICA Y EXCLUSIVAMENTE en los 'Temas técnicos a estudiar'. Si los temas tratan de Bases de Datos, habla de SQL, Modelos E/R, Normalización, etc. NO inventes ejemplos fuera del contexto de la ingeniería de software.\n"
		"2) DETALLE DE ACTIVIDADES EN PÁRRAFOS: En la fila 'Actividades', enumera las acciones paso a paso. Para separar cada paso en un párrafo distinto dentro de la celda de la tabla, DEBES usar obligatoriamente la etiqueta <br><br> entre cada número (ejemplo: 1. Texto paso uno...<br><br>2. Texto paso dos...<br><br>3. Texto paso tres...).\n"
		"3) FORMATO EXACTO: Para CADA semana, usa EXACTAMENTE esta estructura de tabla Markdown de 2 columnas:\n\n"
		"### Semana [X]: [Nombre del tema técnico real a estudiar]\n"
		"| Componente | Descripción Operativa |\n"
		"| :--- | :--- |\n"
		"| **Actividades** | 1. [Instrucción técnica paso 1]<br><br>2. [Instrucción técnica paso 2]<br><br>3. [Instrucción técnica paso 3] |\n"
		"| **Dinámica de Apoyo** | [Cómo interactuar con el tutor/compañeros para resolver dudas técnicas] |\n"
		"| **Coevaluación** | [Actividad de revisión cruzada técnica con compañeros (ej. auditar código o diagramas)] |\n"
		"| **Evidencia Requerida** | [Entregable concreto: ej. Script SQL, Diagrama E/R, Informe técnico] |\n"
		"| **Criterio de Logro** | [Qué define que la tarea técnica fue realizada con éxito] |\n\n"
		"4) NO generes introducciones, saludos ni conclusiones. Tu respuesta debe consistir SOLO en los 4 bloques de semanas con sus tablas.\n"
		"5) OBLIGATORIO: Independientemente de si el alumno reprobó 1 o 5 temas, DEBES generar SIEMPRE una estructura de EXACTAMENTE 4 semanas. Si hay pocos temas, distribúyelos o incluye semanas de repaso y evaluación final. Si hay muchos temas, agrúpalos.\n"
	)

	if colectivo:
		prompt = ChatPromptTemplate.from_messages([
			("system", "Eres un Profesor Militar de Ingeniería. Genera planes de estudio técnicos, detallados, sin alucinaciones bélicas y en formato estricto de tabla."),
			("human", instrucciones_base)
		])
	else:
		prompt = ChatPromptTemplate.from_messages([
			("system", "Eres un Profesor Militar de Ingeniería. Genera planes de estudio técnicos, detallados, sin alucinaciones bélicas y en formato estricto de tabla."),
			("human", instrucciones_base)
		])

	llm = _crear_llm()
	parser = StrOutputParser()

	return RunnableLambda(_armar_entrada_llm) | prompt | llm | parser


def generar_plan_trabajo_alumno(
	nombre_alumno: str,
	tau: float = DEFAULT_TAU,
	brechas_forzadas: list[str] = None,
) -> dict[str, Any]:
	"""
	Ejecuta la cadena y retorna plan junto con metadatos de auditoria basicos.
	"""
	chain = construir_cadena_plan_4_semanas(colectivo=False)
	plan = chain.invoke(
		{
			"nombre_alumno": nombre_alumno,
			"tau": tau,
			"brechas_forzadas": brechas_forzadas,
			"modo_cohorte": False,
		}
	)

	if brechas_forzadas is not None:
		brechas = {
			"alumno": nombre_alumno,
			"puntaje_total": None,
			"nota_final": None,
			"indicadores_brecha": [
				str(item).strip() for item in brechas_forzadas if str(item).strip()
			],
		}
	else:
		brechas = obtener_brechas_alumno(nombre_alumno)
	return {
		"alumno": brechas.get("alumno", nombre_alumno),
		"alumnos_grupo": brechas.get("alumnos_grupo", []),
		"puntaje_total": brechas.get("puntaje_total"),
		"nota_final": brechas.get("nota_final"),
		"indicadores_brecha": brechas.get("indicadores_brecha", []),
		"tau": tau,
		"plan_4_semanas": plan,
	}


def generar_plan_trabajo_cohorte(
	alumnos_grupo: list[str],
	brechas_forzadas: list[str] | None = None,
	tau: float = DEFAULT_TAU,
	unidad_id: str = None,
) -> dict[str, Any]:
	"""Genera un plan colectivo por cohorte manteniendo el mismo formato de salida y conectando el RAG."""
	grupo_limpio = [str(item).strip() for item in alumnos_grupo if str(item).strip()]
	if not grupo_limpio:
		raise ValueError("Debe indicar al menos un alumno en alumnos_grupo.")

	chain = construir_cadena_plan_4_semanas(colectivo=True)
	plan = chain.invoke(
		{
			"nombre_alumno": ", ".join(grupo_limpio),
			"alumnos_grupo": grupo_limpio,
			"tau": tau,
			"brechas_forzadas": brechas_forzadas,
			"modo_cohorte": True,
			"unidad_id": unidad_id,
		}
	)

	if brechas_forzadas is not None:
		brechas = {
			"alumno": ", ".join(grupo_limpio),
			"alumnos_grupo": grupo_limpio,
			"puntaje_total": None,
			"nota_final": None,
			"indicadores_brecha": _normalizar_brechas(brechas_forzadas),
		}
	else:
		brechas = {
			"alumno": ", ".join(grupo_limpio),
			"alumnos_grupo": grupo_limpio,
			"puntaje_total": None,
			"nota_final": None,
			"indicadores_brecha": [],
		}

	return {
		"alumno": brechas.get("alumno", ", ".join(grupo_limpio)),
		"alumnos_grupo": brechas.get("alumnos_grupo", grupo_limpio),
		"puntaje_total": brechas.get("puntaje_total"),
		"nota_final": brechas.get("nota_final"),
		"indicadores_brecha": brechas.get("indicadores_brecha", []),
		"tau": tau,
		"plan_4_semanas": plan,
	}


def generar_sugerencia_pregunta_cobertura(indicador: str, unidad_id: str, tau: float = DEFAULT_TAU) -> str:
	"""Genera una propuesta de pregunta para cubrir un indicador huerfano usando la bibliografia especifica de la unidad."""
	contexto = _recuperar_contexto_por_indicador(indicador=indicador, tau=tau, unidad_id=unidad_id)

	if not contexto.fragmentos:
		return (
			f"⚠️ **Atención:** No se encontró doctrina suficiente (con umbral similitud > {tau}) en la bibliografía "
			f"vectorizada de esta unidad para el indicador {indicador}. Ve a 'Bibliografía y Vectorización' y carga documentos pertinentes."
		)

	doctrina_texto = "\n".join([f"- {frag}" for frag in contexto.fragmentos])

	prompt = ChatPromptTemplate.from_messages([
		("system", "Eres un evaluador académico militar experto en diseño curricular por competencias. Tu misión es formular reactivos técnicos basados estrictamente en la bibliografía aportada."),
		("human",
		 "Genera una propuesta formal de pregunta de opción múltiple para evaluar el indicador: {indicador}.\n\n"
		 "Requisitos:\n"
		 "1) Enunciado claro y contextualizado.\n"
		 "2) Cuatro alternativas (a, b, c, d), con solo una correcta.\n"
		 "3) **Justificación explícita:** Explica por qué la alternativa es correcta, citando y haciendo referencia directa a la doctrina provista a continuación.\n\n"
		 "Doctrina militar extraída de la unidad:\n{contexto_refuerzo}")
	])

	llm = _crear_llm()
	parser = StrOutputParser()
	chain = prompt | llm | parser

	return chain.invoke({"indicador": indicador, "contexto_refuerzo": doctrina_texto})
