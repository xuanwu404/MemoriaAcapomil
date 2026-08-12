from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import os

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


DEFAULT_OLLAMA_MODEL = "gemma4:latest"
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


def _recuperar_contexto_por_indicador(indicador: str, tau: float) -> ContextoRefuerzo:
	consulta = (
		"Refuerzo doctrinario para mejorar el indicador "
		f"{indicador} en evaluacion por competencias."
	)
	resultados = recuperar_fragmentos_relevantes(
		pregunta=consulta,
		tau=tau,
		k=4,
	)
	fragmentos = [item.fragmento for item in resultados]
	return ContextoRefuerzo(indicador=indicador, fragmentos=fragmentos)


def _formatear_contexto_refuerzo(contextos: list[ContextoRefuerzo]) -> str:
	bloques: list[str] = []

	for contexto in contextos:
		if not contexto.fragmentos:
			bloques.append(
				f"Indicador {contexto.indicador}: sin evidencia doctrinaria sobre tau."
			)
			continue

		bloques.append(f"Indicador {contexto.indicador}:")
		for idx, fragmento in enumerate(contexto.fragmentos, start=1):
			bloques.append(f"- Fragmento {idx}: {fragmento}")

	return "\n".join(bloques)


def _normalizar_brechas(brechas_forzadas: list[str] | None) -> list[str]:
	return [str(item).strip() for item in brechas_forzadas or [] if str(item).strip()]


def _armar_entrada_llm(payload: dict[str, Any]) -> dict[str, Any]:
	nombre_alumno = str(payload.get("nombre_alumno", "")).strip()
	alumnos_grupo = [str(item).strip() for item in payload.get("alumnos_grupo", []) if str(item).strip()]
	tau = float(payload.get("tau", DEFAULT_TAU))
	brechas_forzadas = payload.get("brechas_forzadas")
	modo_cohorte = bool(payload.get("modo_cohorte", False))

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

	contextos = [
		_recuperar_contexto_por_indicador(indicador=cod, tau=tau)
		for cod in indicadores
	]

	return {
		"alumno": resumen.get("alumno", nombre_alumno),
		"alumnos_grupo": resumen.get("alumnos_grupo", alumnos_grupo),
		"puntaje_total": resumen.get("puntaje_total"),
		"nota_final": resumen.get("nota_final"),
		"indicadores_brecha": ", ".join(indicadores) if indicadores else "Sin brechas",
		"contexto_refuerzo": _formatear_contexto_refuerzo(contextos),
		"tau": tau,
	}


def construir_cadena_plan_4_semanas(colectivo: bool = False):
	"""Construye una cadena LCEL para generar plan pedagogico individual o colectivo."""
	_asegurar_retrieval_disponible()
	if colectivo:
		prompt = ChatPromptTemplate.from_messages(
			[
				(
					"system",
					"Eres un asesor pedagogico militar experto en didactica por competencias. "
					"Genera planes accionables, claros y medibles para una cohorte de cadetes. "
					"Debes promover apoyo entre pares, tutorias entre cadetes y trabajo colaborativo."
				),
				(
					"human",
					"Genera un Plan de Trabajo Colectivo por Cohorte de 4 semanas para la cohorte {alumno}.\n"
					"Cadetes del grupo:\n"
					"{alumnos_grupo}\n"
					"Datos de la cohorte:\n"
					"- Puntaje total: {puntaje_total}\n"
					"- Nota final: {nota_final}\n"
					"- Indicadores reprobados (<0.60): {indicadores_brecha}\n"
					"- Umbral de confianza retrieval (tau): {tau}\n\n"
					"Contexto doctrinario recuperado:\n"
					"{contexto_refuerzo}\n\n"
					"Requisitos de salida:\n"
					"1) Secciones por Semana 1, 2, 3 y 4.\n"
					"2) Cada semana debe incluir objetivo, actividades, evidencia y criterio de logro.\n"
					"3) Relaciona cada actividad con indicadores reprobados especificos.\n"
					"4) Incluye dinamicas de apoyo entre pares, tutorias entre cadetes y seguimiento del instructor.\n"
					"5) Formula actividades de cohesion de cohorte y coevaluacion.\n"
					"6) Escribe en espanol tecnico, claro y operativo."
				),
			]
		)
	else:
		prompt = ChatPromptTemplate.from_messages(
			[
				(
					"system",
					"Eres un asesor pedagogico militar experto en didactica por competencias. "
					"Genera planes accionables, claros y medibles."
				),
				(
					"human",
					"Genera un plan de trabajo pedagogico detallado de 4 semanas para el alumno {alumno}.\n"
					"Datos del alumno:\n"
					"- Puntaje total: {puntaje_total}\n"
					"- Nota final: {nota_final}\n"
					"- Indicadores reprobados (<0.60): {indicadores_brecha}\n"
					"- Umbral de confianza retrieval (tau): {tau}\n\n"
					"Contexto doctrinario recuperado:\n"
					"{contexto_refuerzo}\n\n"
					"Requisitos de salida:\n"
					"1) Secciones por Semana 1, 2, 3 y 4.\n"
					"2) Cada semana debe incluir objetivo, actividades, evidencia y criterio de logro.\n"
					"3) Relaciona cada actividad con indicadores reprobados especificos.\n"
					"4) Incluir estrategia de seguimiento docente y autoevaluacion del alumno.\n"
					"5) Escribe en espanol tecnico, claro y operativo."
				),
			]
		)

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
) -> dict[str, Any]:
	"""Genera un plan colectivo por cohorte manteniendo el mismo formato de salida."""
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
