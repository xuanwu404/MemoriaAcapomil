from __future__ import annotations

from functools import lru_cache
import importlib
import io
import re
from pathlib import Path
from typing import Any
from datetime import datetime
import sqlite3
import unicodedata
import zipfile

import pandas as pd


UMBRAL_APROBACION = 0.60
ARCHIVO_NOTAS = Path(__file__).resolve().parent.parent / "Archivos Proyecto" / "Notas.xlsx"
HOJA_PUNTAJES = "Puntajes"
HOJA_INDICADORES = "Indicadores de desempeño"
DB_PATH = Path(__file__).resolve().parent.parent / "data" / "acapomil.db"


def _get_connection() -> sqlite3.Connection:
	DB_PATH.parent.mkdir(parents=True, exist_ok=True)
	conn = sqlite3.connect(DB_PATH)
	conn.row_factory = sqlite3.Row
	conn.execute("PRAGMA foreign_keys = ON")
	return conn


def init_database() -> None:
	"""Inicializa esquema SQLite para persistencia relacional del sistema."""
	with _get_connection() as conn:
		conn.executescript(
			"""
			PRAGMA foreign_keys = OFF;

			DROP TABLE IF EXISTS usuarios;
			CREATE TABLE usuarios (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				username TEXT NOT NULL UNIQUE,
				nombre_completo TEXT NOT NULL,
				clave TEXT NOT NULL,
				rol TEXT NOT NULL DEFAULT 'docente' CHECK(rol IN ('docente', 'admin'))
			);

			PRAGMA foreign_keys = ON;

			CREATE TABLE IF NOT EXISTS unidades (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				nombre TEXT NOT NULL,
				anio TEXT NOT NULL,
				mencion TEXT NOT NULL,
				docente_id INTEGER NOT NULL,
				FOREIGN KEY (docente_id) REFERENCES usuarios(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS alumnos (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				rut TEXT NOT NULL,
				nombre TEXT NOT NULL,
				email TEXT NOT NULL DEFAULT '',
				anio TEXT NOT NULL,
				mencion TEXT NOT NULL,
				UNIQUE(rut, anio, mencion)
			);

			CREATE TABLE IF NOT EXISTS resultados (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				alumno_id INTEGER NOT NULL,
				indicador TEXT NOT NULL,
				logro REAL NOT NULL,
				observacion TEXT,
				FOREIGN KEY (alumno_id) REFERENCES alumnos(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS respuestas_alumno (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				alumno_id INTEGER NOT NULL,
				evaluacion_id INTEGER NOT NULL,
				numero_pregunta TEXT NOT NULL,
				puntos_obtenidos REAL NOT NULL,
				UNIQUE(alumno_id, evaluacion_id, numero_pregunta),
				FOREIGN KEY (alumno_id) REFERENCES alumnos(id) ON DELETE CASCADE,
				FOREIGN KEY (evaluacion_id) REFERENCES evaluaciones(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS indicadores_unidad (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				unidad_id INTEGER NOT NULL,
				codigo TEXT NOT NULL,
				descripcion TEXT,
				UNIQUE(unidad_id, codigo),
				FOREIGN KEY (unidad_id) REFERENCES unidades(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS evaluaciones (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				unidad_id INTEGER NOT NULL,
				nombre TEXT NOT NULL,
				archivo_path TEXT,
				FOREIGN KEY (unidad_id) REFERENCES unidades(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS mapeo_preguntas (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				evaluacion_id INTEGER NOT NULL,
				numero_pregunta TEXT NOT NULL,
				puntos_maximos REAL NOT NULL,
				indicadores_asociados TEXT NOT NULL,
				UNIQUE(evaluacion_id, numero_pregunta),
				FOREIGN KEY (evaluacion_id) REFERENCES evaluaciones(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS retroalimentaciones (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				rut_alumno TEXT NOT NULL,
				unidad_id TEXT NOT NULL,
				fecha_generacion TEXT NOT NULL,
				contenido_plan TEXT NOT NULL,
				tau_utilizado REAL NOT NULL,
				FOREIGN KEY (unidad_id) REFERENCES unidades(id) ON DELETE CASCADE
			);
			"""
		)

		_migrar_tabla_alumnos_si_requiere(conn)
		_migrar_tabla_alumnos_email(conn)
		_migrar_tabla_mapeo_preguntas_si_requiere(conn)

		# Xuanwu: semilla fija y unica exigida para un arranque limpio controlado.
		conn.executemany(
			"""
			INSERT INTO usuarios (username, nombre_completo, clave, rol)
			VALUES (?, ?, ?, ?)
			""",
			[
				("admin", "Administrador Principal", "admin123", "admin"),
				("123", "123", "123", "docente"),
			],
		)


def _asegurar_esquema_no_destructivo() -> None:
	"""Asegura tablas y semillas minimas sin eliminar informacion existente."""
	with _get_connection() as conn:
		conn.executescript(
			"""
			CREATE TABLE IF NOT EXISTS usuarios (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				username TEXT NOT NULL UNIQUE,
				nombre_completo TEXT NOT NULL,
				clave TEXT NOT NULL,
				rol TEXT NOT NULL DEFAULT 'docente' CHECK(rol IN ('docente', 'admin'))
			);

			CREATE TABLE IF NOT EXISTS unidades (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				nombre TEXT NOT NULL,
				anio TEXT NOT NULL,
				mencion TEXT NOT NULL,
				docente_id INTEGER NOT NULL,
				FOREIGN KEY (docente_id) REFERENCES usuarios(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS alumnos (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				rut TEXT NOT NULL,
				nombre TEXT NOT NULL,
				email TEXT NOT NULL DEFAULT '',
				anio TEXT NOT NULL,
				mencion TEXT NOT NULL,
				UNIQUE(rut, anio, mencion)
			);

			CREATE TABLE IF NOT EXISTS resultados (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				alumno_id INTEGER NOT NULL,
				indicador TEXT NOT NULL,
				logro REAL NOT NULL,
				observacion TEXT,
				FOREIGN KEY (alumno_id) REFERENCES alumnos(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS respuestas_alumno (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				alumno_id INTEGER NOT NULL,
				evaluacion_id INTEGER NOT NULL,
				numero_pregunta TEXT NOT NULL,
				puntos_obtenidos REAL NOT NULL,
				UNIQUE(alumno_id, evaluacion_id, numero_pregunta),
				FOREIGN KEY (alumno_id) REFERENCES alumnos(id) ON DELETE CASCADE,
				FOREIGN KEY (evaluacion_id) REFERENCES evaluaciones(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS indicadores_unidad (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				unidad_id INTEGER NOT NULL,
				codigo TEXT NOT NULL,
				descripcion TEXT,
				UNIQUE(unidad_id, codigo),
				FOREIGN KEY (unidad_id) REFERENCES unidades(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS evaluaciones (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				unidad_id INTEGER NOT NULL,
				nombre TEXT NOT NULL,
				archivo_path TEXT,
				FOREIGN KEY (unidad_id) REFERENCES unidades(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS mapeo_preguntas (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				evaluacion_id INTEGER NOT NULL,
				numero_pregunta TEXT NOT NULL,
				puntos_maximos REAL NOT NULL,
				indicadores_asociados TEXT NOT NULL,
				UNIQUE(evaluacion_id, numero_pregunta),
				FOREIGN KEY (evaluacion_id) REFERENCES evaluaciones(id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS retroalimentaciones (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				rut_alumno TEXT NOT NULL,
				unidad_id TEXT NOT NULL,
				fecha_generacion TEXT NOT NULL,
				contenido_plan TEXT NOT NULL,
				tau_utilizado REAL NOT NULL,
				FOREIGN KEY (unidad_id) REFERENCES unidades(id) ON DELETE CASCADE
			);
			"""
		)

		_migrar_tabla_alumnos_si_requiere(conn)
		_migrar_tabla_alumnos_email(conn)
		_migrar_tabla_mapeo_preguntas_si_requiere(conn)

		# Xuanwu: semilla fija y unica exigida para un arranque limpio controlado.
		conn.execute(
			"""
			INSERT INTO usuarios (username, nombre_completo, clave, rol)
			SELECT ?, ?, ?, ?
			WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE username = ?)
			""",
			("admin", "Administrador Principal", "admin123", "admin", "admin"),
		)
		conn.execute(
			"""
			INSERT INTO usuarios (username, nombre_completo, clave, rol)
			SELECT ?, ?, ?, ?
			WHERE NOT EXISTS (SELECT 1 FROM usuarios WHERE username = ?)
			""",
			("123", "123", "123", "docente", "123"),
		)


def _limpiar_texto_pdf(texto: str) -> str:
	texto_limpio = unicodedata.normalize("NFKD", str(texto or ""))
	texto_limpio = "".join(c for c in texto_limpio if not unicodedata.combining(c))
	texto_limpio = re.sub(r"\s+", " ", texto_limpio)
	return texto_limpio.strip()


def _es_encabezado_pdf(linea: str) -> bool:
	texto = _limpiar_texto_pdf(linea).lower()
	if not texto:
		return True
	return any(
		fragmento in texto
		for fragmento in [
			"aprendizaje esperado por evaluar",
			"tareas asociadas",
			"requerimientos",
			"indicadores de desempeno",
			"unidad de aprendizaje",
			"ejercito de chile",
			"division educacion",
		]
	)


def _limpiar_descripcion_pdf(texto: str) -> str:
	descripcion = _limpiar_texto_pdf(texto)
	for frase in [
		"Aplica los conceptos de integridad y vistas.",
		"Aplica normalizaciones de",
		"Aplica transacciones en bases de datos y los conceptos de seguridad, optimizacion y e informacion",
		"APRENDIZAJE ESPERADO POR EVALUAR TAREAS DESEMPENO",
		"Aplica el diseno conceptual de la base de datos en un",
		"UNIDAD DE APRENDIZAJE: \"BASE DE DATOS SAS\"",
	]:
		descripcion = descripcion.replace(frase, " ")
	descripcion = descripcion.lstrip(". ")
	descripcion = re.sub(r"^[\s\.-:;•·,]+", "", descripcion)
	descripcion = re.sub(r"\s+", " ", descripcion).strip()
	return descripcion


def extraer_indicadores_desde_pdf(pdf_bytes: bytes) -> list[dict[str, str]]:
	"""Extrae codigos y descripciones de indicadores desde un PDF."""
	texto_pdf = ""
	try:
		pypdf_module = importlib.import_module("pypdf")
		reader = pypdf_module.PdfReader(io.BytesIO(pdf_bytes))
		partes: list[str] = []
		for pagina in reader.pages:
			partes.append(pagina.extract_text() or "")
		texto_pdf = "\n".join(partes)
	except Exception:
		try:
			pypdf_module = importlib.import_module("PyPDF2")
			reader = pypdf_module.PdfReader(io.BytesIO(pdf_bytes))
			partes = []
			for pagina in reader.pages:
				partes.append(pagina.extract_text() or "")
			texto_pdf = "\n".join(partes)
		except Exception:
			texto_pdf = ""

	if not texto_pdf:
		return []

	patron_codigo_inicio = re.compile(
		r"(?:^|\s)(?P<codigo>1\.[1-3]|2\.[1-3]|3\.[1-3]|4\.[1-5]|5\.[1-9])\b",
		re.IGNORECASE,
	)
	codigos_esperados = [
		"1.1", "1.2", "1.3",
		"2.1", "2.2", "2.3",
		"3.1", "3.2", "3.3",
		"4.1", "4.2", "4.3", "4.4", "4.5",
		"5.1", "5.2", "5.3", "5.4", "5.5", "5.6", "5.7", "5.8", "5.9",
	]
	mapa_orden = {codigo: indice for indice, codigo in enumerate(codigos_esperados)}
	lineas: list[str] = []
	for bloque in texto_pdf.splitlines():
		linea = _limpiar_texto_pdf(bloque)
		if not linea or _es_encabezado_pdf(linea):
			continue
		for frase in [
			"Aplica los conceptos de integridad y vistas.",
			"Aplica normalizaciones de",
			"Aplica transacciones en bases de datos y los conceptos de seguridad, optimizacion y e informacion",
			"APRENDIZAJE ESPERADO POR EVALUAR TAREAS DESEMPENO",
			"Aplica el diseno conceptual de la base de datos en un",
			"UNIDAD DE APRENDIZAJE: \"BASE DE DATOS SAS\"",
		]:
			linea = linea.replace(frase, " ")
		linea = re.sub(r"(?<=\d)\s*\.\s*(?=\d)", ".", linea)
		linea = re.sub(r"(?<=\b\d)\s+(?=\d\.)", "", linea)
		lineas.append(linea)

	indicadores: list[dict[str, str]] = []
	vistos: set[str] = set()
	codigo_actual: str | None = None
	descripcion_actual: list[str] = []
	buffer_desc: list[str] = []

	def _guardar_actual() -> None:
		nonlocal codigo_actual, descripcion_actual, buffer_desc
		if not codigo_actual:
			return
		descripcion = " ".join(descripcion_actual + buffer_desc)
		descripcion = _limpiar_descripcion_pdf(descripcion)
		if not descripcion:
			codigo_actual = None
			descripcion_actual = []
			buffer_desc = []
			return
		indicadores.append({"codigo": codigo_actual, "descripcion": descripcion})
		vistos.add(codigo_actual)
		codigo_actual = None
		descripcion_actual = []
		buffer_desc = []

	for linea in lineas:
		linea_trabajo = linea.replace("|", " ")
		coincidencia = patron_codigo_inicio.search(linea_trabajo)
		if coincidencia:
			codigo = coincidencia.group("codigo").strip().upper()
			if codigo not in mapa_orden:
				continue
			if codigo_actual and codigo != codigo_actual:
				_guardar_actual()
			if codigo in vistos:
				codigo_actual = codigo
				descripcion_actual = []
				buffer_desc = []
				continue

			codigo_actual = codigo
			descripcion_actual = []
			buffer_desc = []
			resto = linea_trabajo[coincidencia.end():]
			resto = _limpiar_descripcion_pdf(resto)
			if resto:
				buffer_desc.append(resto)
			continue

		match = patron_codigo_inicio.match(linea)
		if codigo_actual:
			buffer_desc.append(_limpiar_descripcion_pdf(linea))

	_guardar_actual()

	indicadores_unicos: list[dict[str, str]] = []
	vistos_finales: set[str] = set()
	for fila in indicadores:
		codigo = str(fila.get("codigo", "")).strip().upper()
		descripcion = _limpiar_descripcion_pdf(fila.get("descripcion", ""))
		if not codigo or codigo not in mapa_orden or codigo in vistos_finales:
			continue
		if descripcion:
			indicadores_unicos.append({"codigo": codigo, "descripcion": descripcion})
			vistos_finales.add(codigo)

	indicadores_unicos.sort(key=lambda fila: mapa_orden.get(fila["codigo"], 999))
	return indicadores_unicos


def inicializar_si_no_existe() -> None:
	"""Inicializa la base de datos solo cuando el archivo fisico aun no existe."""
	if DB_PATH.exists():
		_asegurar_esquema_no_destructivo()
		_asegurar_tabla_respuestas_alumno()
		return
	init_database()
	_asegurar_tabla_respuestas_alumno()


def _asegurar_tabla_respuestas_alumno() -> None:
	with _get_connection() as conn:
		conn.execute(
			"""
			CREATE TABLE IF NOT EXISTS respuestas_alumno (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				alumno_id INTEGER NOT NULL,
				evaluacion_id INTEGER NOT NULL,
				numero_pregunta TEXT NOT NULL,
				puntos_obtenidos REAL NOT NULL,
				UNIQUE(alumno_id, evaluacion_id, numero_pregunta),
				FOREIGN KEY (alumno_id) REFERENCES alumnos(id) ON DELETE CASCADE,
				FOREIGN KEY (evaluacion_id) REFERENCES evaluaciones(id) ON DELETE CASCADE
			)
			"""
		)


def _migrar_tabla_usuarios_si_requiere(conn: sqlite3.Connection) -> None:
	"""Migra tabla usuarios legacy a esquema con username + nombre_completo."""
	columnas = {row["name"] for row in conn.execute("PRAGMA table_info(usuarios)").fetchall()}
	requiere_migracion = "username" not in columnas or "nombre_completo" not in columnas
	if not requiere_migracion and "rol" in columnas:
		return

	conn.execute("PRAGMA foreign_keys = OFF")
	conn.execute(
		"""
		CREATE TABLE IF NOT EXISTS usuarios_new (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			username TEXT NOT NULL UNIQUE,
			nombre_completo TEXT NOT NULL,
			clave TEXT NOT NULL,
			rol TEXT NOT NULL DEFAULT 'docente'
		)
		"""
	)

	rows = conn.execute("SELECT * FROM usuarios ORDER BY id").fetchall()
	for row in rows:
		row_dict = dict(row)
		legacy_nombre = str(row_dict.get("nombre", "")).strip()
		username = str(row_dict.get("username", legacy_nombre)).strip()
		nombre_completo = str(row_dict.get("nombre_completo", legacy_nombre or username)).strip()
		clave = str(row_dict.get("clave", "")).strip()
		rol = str(row_dict.get("rol", "docente")).strip() or "docente"
		if not username:
			username = f"usuario_{row_dict.get('id', '')}".strip("_")
		if not nombre_completo:
			nombre_completo = username
		conn.execute(
			"""
			INSERT OR REPLACE INTO usuarios_new (id, username, nombre_completo, clave, rol)
			VALUES (?, ?, ?, ?, ?)
			""",
			(int(row_dict.get("id", 0)), username, nombre_completo, clave, rol),
		)

	conn.execute("DROP TABLE IF EXISTS usuarios")
	conn.execute("ALTER TABLE usuarios_new RENAME TO usuarios")
	conn.execute("PRAGMA foreign_keys = ON")


def _migrar_tabla_alumnos_email(conn: sqlite3.Connection) -> None:
	columnas = {row["name"] for row in conn.execute("PRAGMA table_info(alumnos)").fetchall()}
	if "email" not in columnas:
		conn.execute("ALTER TABLE alumnos ADD COLUMN email TEXT NOT NULL DEFAULT ''")


def _migrar_tabla_alumnos_si_requiere(conn: sqlite3.Connection) -> None:
	"""Migra esquema antiguo de alumnos (con unidad_id) al nuevo modelo centralizado."""
	columnas = {row["name"] for row in conn.execute("PRAGMA table_info(alumnos)").fetchall()}
	if "unidad_id" not in columnas:
		return

	conn.execute("PRAGMA foreign_keys = OFF")
	conn.execute(
		"""
		CREATE TABLE IF NOT EXISTS alumnos_new (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			rut TEXT NOT NULL,
			nombre TEXT NOT NULL,
			anio TEXT NOT NULL,
			mencion TEXT NOT NULL,
			UNIQUE(rut, anio, mencion)
		)
		"""
	)
	conn.execute(
		"""
		CREATE TABLE IF NOT EXISTS resultados_new (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			alumno_id INTEGER NOT NULL,
			indicador TEXT NOT NULL,
			logro REAL NOT NULL,
			observacion TEXT,
			FOREIGN KEY (alumno_id) REFERENCES alumnos_new(id) ON DELETE CASCADE
		)
		"""
	)

	mapeo_ids: dict[int, int] = {}
	alumnos_old = conn.execute(
		"SELECT id, rut, nombre, anio, mencion FROM alumnos ORDER BY id"
	).fetchall()
	for alumno in alumnos_old:
		row_existente = conn.execute(
			"SELECT id FROM alumnos_new WHERE rut = ? AND anio = ? AND mencion = ?",
			(alumno["rut"], alumno["anio"], alumno["mencion"]),
		).fetchone()
		if row_existente:
			nuevo_id = int(row_existente["id"])
		else:
			cursor = conn.execute(
				"""
				INSERT INTO alumnos_new (rut, nombre, anio, mencion)
				VALUES (?, ?, ?, ?)
				""",
				(alumno["rut"], alumno["nombre"], alumno["anio"], alumno["mencion"]),
			)
			nuevo_id = int(cursor.lastrowid)
		mapeo_ids[int(alumno["id"])] = nuevo_id

	resultados_old = conn.execute(
		"SELECT id, alumno_id, indicador, logro, observacion FROM resultados ORDER BY id"
	).fetchall()
	for resultado in resultados_old:
		nuevo_alumno_id = mapeo_ids.get(int(resultado["alumno_id"]))
		if nuevo_alumno_id is None:
			continue
		conn.execute(
			"""
			INSERT INTO resultados_new (alumno_id, indicador, logro, observacion)
			VALUES (?, ?, ?, ?)
			""",
			(
				nuevo_alumno_id,
				resultado["indicador"],
				resultado["logro"],
				resultado["observacion"],
			),
		)

	conn.execute("DROP TABLE IF EXISTS resultados")
	conn.execute("DROP TABLE IF EXISTS alumnos")
	conn.execute("ALTER TABLE alumnos_new RENAME TO alumnos")
	conn.execute("ALTER TABLE resultados_new RENAME TO resultados")
	conn.execute("PRAGMA foreign_keys = ON")


def _migrar_tabla_mapeo_preguntas_si_requiere(conn: sqlite3.Connection) -> None:
	"""Migra mapeo_preguntas legacy (indicador_asociado) a indicadores_asociados consolidado."""
	columnas = {row["name"] for row in conn.execute("PRAGMA table_info(mapeo_preguntas)").fetchall()}
	if not columnas:
		return
	if "indicadores_asociados" in columnas:
		return

	conn.execute("PRAGMA foreign_keys = OFF")
	conn.execute(
		"""
		CREATE TABLE IF NOT EXISTS mapeo_preguntas_new (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			evaluacion_id INTEGER NOT NULL,
			numero_pregunta TEXT NOT NULL,
			puntos_maximos REAL NOT NULL,
			indicadores_asociados TEXT NOT NULL,
			UNIQUE(evaluacion_id, numero_pregunta),
			FOREIGN KEY (evaluacion_id) REFERENCES evaluaciones(id) ON DELETE CASCADE
		)
		"""
	)

	rows = conn.execute(
		"""
		SELECT id, evaluacion_id, numero_pregunta, puntos_maximos, indicador_asociado
		FROM mapeo_preguntas
		ORDER BY id
		"""
	).fetchall()
	for row in rows:
		indicadores_asociados = str(row["indicador_asociado"] or "").strip()
		if not indicadores_asociados:
			continue
		conn.execute(
			"""
			INSERT OR REPLACE INTO mapeo_preguntas_new
			(id, evaluacion_id, numero_pregunta, puntos_maximos, indicadores_asociados)
			VALUES (?, ?, ?, ?, ?)
			""",
			(
				int(row["id"]),
				int(row["evaluacion_id"]),
				str(row["numero_pregunta"]),
				float(row["puntos_maximos"]),
				indicadores_asociados,
			),
		)

	conn.execute("DROP TABLE IF EXISTS mapeo_preguntas")
	conn.execute("ALTER TABLE mapeo_preguntas_new RENAME TO mapeo_preguntas")
	conn.execute("PRAGMA foreign_keys = ON")


def _row_to_dict(row: sqlite3.Row | None) -> dict[str, Any] | None:
	if row is None:
		return None
	return dict(row)


def _limpiar_rut(rut_str: str) -> str:
	return "".join(c for c in str(rut_str).strip().upper() if c.isalnum())


def validar_rut_chileno(rut_str: str) -> bool:
	"""Valida RUT chileno usando algoritmo modulo 11."""
	rut_limpio = _limpiar_rut(rut_str)
	if len(rut_limpio) < 2:
		return False

	cuerpo = rut_limpio[:-1]
	dv = rut_limpio[-1]
	if not cuerpo.isdigit() or not (dv.isdigit() or dv == "K"):
		return False

	suma = 0
	factor = 2
	for digito in reversed(cuerpo):
		suma += int(digito) * factor
		factor = 2 if factor == 7 else factor + 1

	resto = 11 - (suma % 11)
	if resto == 11:
		dv_esperado = "0"
	elif resto == 10:
		dv_esperado = "K"
	else:
		dv_esperado = str(resto)

	return dv == dv_esperado


def formatear_rut_chileno(rut_str: str) -> str:
	"""Formatea RUT como X.XXX.XXX-X o XX.XXX.XXX-X."""
	rut_limpio = _limpiar_rut(rut_str)
	if len(rut_limpio) < 2:
		return rut_limpio

	cuerpo = rut_limpio[:-1]
	dv = rut_limpio[-1]
	if not cuerpo.isdigit():
		return rut_limpio

	cuerpo = str(int(cuerpo))
	grupos: list[str] = []
	while cuerpo:
		grupos.append(cuerpo[-3:])
		cuerpo = cuerpo[:-3]
	cuerpo_formateado = ".".join(reversed(grupos))

	return f"{cuerpo_formateado}-{dv}"


def crear_usuario(username: str, nombre_completo: str, clave: str, rol: str) -> int:
	rol_normalizado = _normalizar_rol(rol)
	with _get_connection() as conn:
		cursor = conn.execute(
			"INSERT INTO usuarios (username, nombre_completo, clave, rol) VALUES (?, ?, ?, ?)",
			(username.strip(), nombre_completo.strip(), clave, rol_normalizado),
		)
		return int(cursor.lastrowid)


def listar_usuarios() -> list[dict[str, Any]]:
	with _get_connection() as conn:
		rows = conn.execute(
			"SELECT id, username, nombre_completo, clave, rol FROM usuarios ORDER BY username"
		).fetchall()
		return [dict(row) for row in rows]


def obtener_usuarios() -> list[dict[str, Any]]:
	"""Alias semantico para listar usuarios desde capa de aplicacion."""
	return listar_usuarios()


def obtener_usuario_por_username(username: str) -> dict[str, Any] | None:
	with _get_connection() as conn:
		row = conn.execute(
			"SELECT id, username, nombre_completo, clave, rol FROM usuarios WHERE username = ?",
			(username.strip(),),
		).fetchone()
		return _row_to_dict(row)


def obtener_usuario_por_nombre(nombre: str) -> dict[str, Any] | None:
	"""Compatibilidad hacia atras: nombre ahora se interpreta como username."""
	return obtener_usuario_por_username(nombre)


def obtener_usuario_por_credenciales(username: str, clave: str) -> dict[str, Any] | None:
	with _get_connection() as conn:
		row = conn.execute(
			"""
			SELECT id, username, nombre_completo, clave, rol
			FROM usuarios
			WHERE username = ? AND clave = ?
			""",
			(username.strip(), clave),
		).fetchone()
		return _row_to_dict(row)


def actualizar_usuario(
	usuario_id: int,
	username: str | None = None,
	nombre_completo: str | None = None,
	clave: str | None = None,
	rol: str | None = None,
) -> None:
	campos: list[str] = []
	valores: list[Any] = []
	rol_normalizado = _normalizar_rol(rol) if rol is not None else None

	with _get_connection() as conn:
		usuario_actual = conn.execute(
			"SELECT username, rol FROM usuarios WHERE id = ?",
			(usuario_id,),
		).fetchone()
		if usuario_actual is None:
			return

		if (
			usuario_actual["username"] == "admin"
			and rol_normalizado is not None
			and rol_normalizado != "admin"
		):
			raise ValueError("La cuenta admin no puede cambiar su rol a docente.")

		if username is not None:
			campos.append("username = ?")
			valores.append(username.strip())
		if nombre_completo is not None:
			campos.append("nombre_completo = ?")
			valores.append(nombre_completo.strip())
		if clave is not None:
			campos.append("clave = ?")
			valores.append(clave)
		if rol_normalizado is not None:
			campos.append("rol = ?")
			valores.append(rol_normalizado)

		if not campos:
			return

		valores.append(usuario_id)
		conn.execute(
			f"UPDATE usuarios SET {', '.join(campos)} WHERE id = ?",
			tuple(valores),
		)


def eliminar_usuario(usuario_id: int) -> None:
	with _get_connection() as conn:
		row = conn.execute("SELECT username FROM usuarios WHERE id = ?", (usuario_id,)).fetchone()
		if row and row["username"] == "admin":
			raise ValueError("La cuenta admin no puede eliminarse.")
		unidades_asignadas = conn.execute(
			"SELECT 1 FROM unidades WHERE docente_id = ? LIMIT 1",
			(usuario_id,),
		).fetchone()
		if unidades_asignadas:
			raise sqlite3.IntegrityError("docente_con_unidades")
		conn.execute("DELETE FROM usuarios WHERE id = ?", (usuario_id,))


def eliminar_usuario_por_username(username: str) -> None:
	with _get_connection() as conn:
		row = conn.execute(
			"SELECT id, username FROM usuarios WHERE username = ?",
			(username.strip(),),
		).fetchone()
		if row is None:
			return
		if row["username"] == "admin":
			raise ValueError("La cuenta admin no puede eliminarse.")
		unidades_asignadas = conn.execute(
			"SELECT 1 FROM unidades WHERE docente_id = ? LIMIT 1",
			(int(row["id"]),),
		).fetchone()
		if unidades_asignadas:
			raise sqlite3.IntegrityError("docente_con_unidades")
		conn.execute("DELETE FROM usuarios WHERE username = ?", (username.strip(),))


def _normalizar_rol(rol: str | None) -> str:
	rol_limpio = str(rol or "").strip().lower()
	if rol_limpio not in {"docente", "admin"}:
		raise ValueError("Rol invalido. Valores permitidos: 'docente' o 'admin'.")
	return rol_limpio


def crear_unidad(nombre: str, anio: str, mencion: str, docente_id: int) -> int:
	with _get_connection() as conn:
		cursor = conn.execute(
			"""
			INSERT INTO unidades (nombre, anio, mencion, docente_id)
			VALUES (?, ?, ?, ?)
			""",
			(nombre.strip(), anio, mencion, docente_id),
		)
		return int(cursor.lastrowid)


def listar_unidades(docente_id: int | None = None) -> list[dict[str, Any]]:
	with _get_connection() as conn:
		if docente_id is None:
			rows = conn.execute(
				"SELECT id, nombre, anio, mencion, docente_id FROM unidades ORDER BY id DESC"
			).fetchall()
		else:
			rows = conn.execute(
				"""
				SELECT id, nombre, anio, mencion, docente_id
				FROM unidades
				WHERE docente_id = ?
				ORDER BY id DESC
				""",
				(docente_id,),
			).fetchall()
		return [dict(row) for row in rows]


def obtener_unidad_por_id(unidad_id: int) -> dict[str, Any] | None:
	with _get_connection() as conn:
		row = conn.execute(
			"SELECT id, nombre, anio, mencion, docente_id FROM unidades WHERE id = ?",
			(unidad_id,),
		).fetchone()
		return _row_to_dict(row)


def actualizar_unidad(unidad_id: int, nombre: str, anio: str, mencion: str) -> None:
	with _get_connection() as conn:
		conn.execute(
			"""
			UPDATE unidades
			SET nombre = ?, anio = ?, mencion = ?
			WHERE id = ?
			""",
			(nombre.strip(), anio, mencion, unidad_id),
		)


def eliminar_unidad(unidad_id: int) -> None:
	with _get_connection() as conn:
		conn.execute("DELETE FROM unidades WHERE id = ?", (unidad_id,))


def crear_alumno(rut: str, nombre: str, email: str, anio: str, mencion: str) -> int:
	if not validar_rut_chileno(rut):
		raise ValueError("RUT invalido segun modulo 11.")
	rut_formateado = formatear_rut_chileno(rut)
	with _get_connection() as conn:
		cursor = conn.execute(
			"""
			INSERT INTO alumnos (rut, nombre, email, anio, mencion)
			VALUES (?, ?, ?, ?, ?)
			""",
			(rut_formateado, nombre.strip(), email.strip(), anio, mencion),
		)
		return int(cursor.lastrowid)


def listar_alumnos(anio: str | None = None, mencion: str | None = None) -> list[dict[str, Any]]:
	with _get_connection() as conn:
		query = "SELECT id, rut, nombre, email, anio, mencion FROM alumnos"
		params: list[Any] = []
		filtros: list[str] = []
		if anio is not None:
			filtros.append("anio = ?")
			params.append(anio)
		if mencion is not None:
			filtros.append("mencion = ?")
			params.append(mencion)
		if filtros:
			query += " WHERE " + " AND ".join(filtros)
		query += " ORDER BY nombre"
		rows = conn.execute(query, tuple(params)).fetchall()
		return [dict(row) for row in rows]


def buscar_alumno_por_rut_anio_mencion(rut: str, anio: str, mencion: str) -> dict[str, Any] | None:
	rut_busqueda = formatear_rut_chileno(rut) if validar_rut_chileno(rut) else str(rut).strip()
	with _get_connection() as conn:
		row = conn.execute(
			"""
			SELECT id, rut, nombre, email, anio, mencion
			FROM alumnos
			WHERE rut = ? AND anio = ? AND mencion = ?
			""",
			(rut_busqueda, anio, mencion),
		).fetchone()
		return _row_to_dict(row)


def eliminar_alumno(alumno_id: int) -> None:
	with _get_connection() as conn:
		conn.execute("DELETE FROM alumnos WHERE id = ?", (alumno_id,))


def reemplazar_resultados_alumno(alumno_id: int, resultados: list[dict[str, Any]]) -> None:
	with _get_connection() as conn:
		conn.execute("DELETE FROM resultados WHERE alumno_id = ?", (alumno_id,))
		for fila in resultados:
			indicador = str(fila.get("Indicador", "")).strip()
			if not indicador:
				continue
			try:
				logro = float(fila.get("Logro", 0))
			except (TypeError, ValueError):
				logro = 0.0
			observacion = str(fila.get("Observacion", "")).strip()
			conn.execute(
				"""
				INSERT INTO resultados (alumno_id, indicador, logro, observacion)
				VALUES (?, ?, ?, ?)
				""",
				(alumno_id, indicador, logro, observacion),
			)


def listar_resultados_alumno(alumno_id: int) -> list[dict[str, Any]]:
	with _get_connection() as conn:
		rows = conn.execute(
			"""
			SELECT id, alumno_id, indicador, logro, observacion
			FROM resultados
			WHERE alumno_id = ?
			ORDER BY id
			""",
			(alumno_id,),
		).fetchall()
		return [dict(row) for row in rows]


def eliminar_resultado(resultado_id: int) -> None:
	with _get_connection() as conn:
		conn.execute("DELETE FROM resultados WHERE id = ?", (resultado_id,))


def eliminar_resultados_alumno(alumno_id: int) -> None:
	with _get_connection() as conn:
		conn.execute("DELETE FROM resultados WHERE alumno_id = ?", (alumno_id,))


def crear_indicador_unidad(unidad_id: int, codigo: str, descripcion: str) -> int:
	with _get_connection() as conn:
		cursor = conn.execute(
			"""
			INSERT INTO indicadores_unidad (unidad_id, codigo, descripcion)
			VALUES (?, ?, ?)
			""",
			(unidad_id, codigo.strip(), descripcion.strip()),
		)
		return int(cursor.lastrowid)


def listar_indicadores_unidad(unidad_id: int) -> list[dict[str, Any]]:
	with _get_connection() as conn:
		rows = conn.execute(
			"""
			SELECT id, unidad_id, codigo, descripcion
			FROM indicadores_unidad
			WHERE unidad_id = ?
			ORDER BY codigo
			""",
			(unidad_id,),
		).fetchall()
		return [dict(row) for row in rows]


def reemplazar_indicadores_unidad(unidad_id: int, filas: list[dict[str, Any]]) -> None:
	with _get_connection() as conn:
		conn.execute("DELETE FROM indicadores_unidad WHERE unidad_id = ?", (unidad_id,))
		for fila in filas:
			codigo = str(fila.get("codigo", "")).strip()
			descripcion = str(fila.get("descripcion", "")).strip()
			if not codigo:
				continue
			conn.execute(
				"""
				INSERT INTO indicadores_unidad (unidad_id, codigo, descripcion)
				VALUES (?, ?, ?)
				""",
				(unidad_id, codigo, descripcion),
			)


def cargar_indicadores_desde_df(unidad_id: int, df_indicadores: pd.DataFrame) -> None:
	"""Carga masiva de indicadores desde DataFrame (Codigo_Completo, Descripcion)."""
	with _get_connection() as conn:
		conn.execute("BEGIN")
		try:
			conn.execute("DELETE FROM indicadores_unidad WHERE unidad_id = ?", (unidad_id,))
			for _, fila in df_indicadores.iterrows():
				codigo = str(fila.get("Codigo_Completo", "")).strip()
				descripcion = str(fila.get("Descripcion", "")).strip()
				if codigo.lower() == "nan":
					codigo = ""
				if descripcion.lower() == "nan":
					descripcion = ""
				if not codigo:
					continue
				conn.execute(
					"""
					INSERT INTO indicadores_unidad (unidad_id, codigo, descripcion)
					VALUES (?, ?, ?)
					""",
					(unidad_id, codigo, descripcion),
				)
			conn.commit()
		except Exception:
			conn.rollback()
			raise


def guardar_evaluacion(unidad_id: int, nombre: str, archivo_path: str) -> int:
	with _get_connection() as conn:
		cursor = conn.execute(
			"""
			INSERT INTO evaluaciones (unidad_id, nombre, archivo_path)
			VALUES (?, ?, ?)
			""",
			(unidad_id, nombre.strip(), archivo_path.strip()),
		)
		return int(cursor.lastrowid)


def listar_evaluaciones_unidad(unidad_id: int) -> list[dict[str, Any]]:
	with _get_connection() as conn:
		rows = conn.execute(
			"""
			SELECT id, unidad_id, nombre, archivo_path
			FROM evaluaciones
			WHERE unidad_id = ?
			ORDER BY id DESC
			""",
			(unidad_id,),
		).fetchall()
		return [dict(row) for row in rows]


def eliminar_evaluacion(evaluacion_id: int) -> None:
	with _get_connection() as conn:
		conn.execute("DELETE FROM evaluaciones WHERE id = ?", (evaluacion_id,))


def guardar_mapeo_pregunta(
	evaluacion_id: int,
	numero_pregunta: str,
	puntos_maximos: float,
	indicadores_asociados: str,
) -> int:
	with _get_connection() as conn:
		cursor = conn.execute(
			"""
			INSERT INTO mapeo_preguntas (evaluacion_id, numero_pregunta, puntos_maximos, indicadores_asociados)
			VALUES (?, ?, ?, ?)
			""",
			(
				evaluacion_id,
				numero_pregunta.strip(),
				float(puntos_maximos),
				indicadores_asociados.strip(),
			),
		)
		return int(cursor.lastrowid)


def listar_mapeo_evaluacion(evaluacion_id: int) -> list[dict[str, Any]]:
	with _get_connection() as conn:
		columnas = {row["name"] for row in conn.execute("PRAGMA table_info(mapeo_preguntas)").fetchall()}
		if "indicadores_asociados" in columnas:
			rows = conn.execute(
				"""
				SELECT id, evaluacion_id, numero_pregunta, puntos_maximos, indicadores_asociados
				FROM mapeo_preguntas
				WHERE evaluacion_id = ?
				ORDER BY numero_pregunta
				""",
				(evaluacion_id,),
			).fetchall()
			return [dict(row) for row in rows]

		rows = conn.execute(
			"""
			SELECT id, evaluacion_id, numero_pregunta, puntos_maximos, indicador_asociado AS indicadores_asociados
			FROM mapeo_preguntas
			WHERE evaluacion_id = ?
			ORDER BY numero_pregunta
			""",
			(evaluacion_id,),
		).fetchall()
		return [dict(row) for row in rows]


def reemplazar_mapeo_evaluacion(evaluacion_id: int, filas: list[dict[str, Any]]) -> None:
	with _get_connection() as conn:
		conn.execute("DELETE FROM mapeo_preguntas WHERE evaluacion_id = ?", (evaluacion_id,))
		for fila in filas:
			numero_pregunta = str(fila.get("numero_pregunta", "")).strip()
			indicadores_asociados = str(
				fila.get("indicadores_asociados", fila.get("indicador_asociado", ""))
			).strip()
			if not numero_pregunta or not indicadores_asociados:
				continue
			try:
				puntos_maximos = float(fila.get("puntos_maximos", 0))
			except (TypeError, ValueError):
				puntos_maximos = 0.0
			if puntos_maximos <= 0:
				continue
			conn.execute(
				"""
				INSERT INTO mapeo_preguntas (evaluacion_id, numero_pregunta, puntos_maximos, indicadores_asociados)
				VALUES (?, ?, ?, ?)
				""",
				(evaluacion_id, numero_pregunta, puntos_maximos, indicadores_asociados),
			)


def guardar_puntajes_pregunta(alumno_id: int, evaluacion_id: int, puntajes: list[dict[str, Any]]) -> None:
	with _get_connection() as conn:
		for fila in puntajes:
			numero_pregunta = str(fila.get("numero_pregunta", "")).strip()
			if not numero_pregunta:
				continue
			try:
				puntos_obtenidos = float(fila.get("puntos_obtenidos", 0.0))
			except (TypeError, ValueError):
				puntos_obtenidos = 0.0
			conn.execute(
				"""
				INSERT OR REPLACE INTO respuestas_alumno (
					alumno_id,
					evaluacion_id,
					numero_pregunta,
					puntos_obtenidos
				)
				VALUES (?, ?, ?, ?)
				""",
				(int(alumno_id), int(evaluacion_id), numero_pregunta, puntos_obtenidos),
			)


def obtener_puntajes_pregunta(alumno_id: int, evaluacion_id: int) -> dict[str, float]:
	with _get_connection() as conn:
		rows = conn.execute(
			"""
			SELECT numero_pregunta, puntos_obtenidos
			FROM respuestas_alumno
			WHERE alumno_id = ? AND evaluacion_id = ?
			""",
			(int(alumno_id), int(evaluacion_id)),
		).fetchall()
		return {str(row["numero_pregunta"]): float(row["puntos_obtenidos"]) for row in rows}


def guardar_retroalimentacion_local(
	rut_alumno: str,
	unidad_id: str,
	plan_texto: str,
	tau: float,
) -> bool:
	if not str(rut_alumno).strip() or not str(unidad_id).strip() or not str(plan_texto).strip():
		return False

	try:
		with _get_connection() as conn:
			conn.execute(
				"""
				INSERT INTO retroalimentaciones (rut_alumno, unidad_id, fecha_generacion, contenido_plan, tau_utilizado)
				VALUES (?, ?, ?, ?, ?)
				""",
				(
					str(rut_alumno).strip(),
					str(unidad_id).strip(),
					datetime.now().isoformat(timespec="seconds"),
					str(plan_texto),
					float(tau),
				),
			)
		return True
	except Exception:
		return False


def obtener_historial_retroalimentaciones(rut_alumno: str, unidad_id: str) -> list[dict[str, Any]]:
	with _get_connection() as conn:
		rows = conn.execute(
			"""
			SELECT id, rut_alumno, unidad_id, fecha_generacion, contenido_plan, tau_utilizado
			FROM retroalimentaciones
			WHERE rut_alumno = ? AND unidad_id = ?
			ORDER BY fecha_generacion DESC, id DESC
			""",
			(str(rut_alumno).strip(), str(unidad_id).strip()),
		).fetchall()
		return [dict(row) for row in rows]


def eliminar_retroalimentacion_local(retro_id: int) -> bool:
	try:
		with _get_connection() as conn:
			conn.execute("DELETE FROM retroalimentaciones WHERE id = ?", (retro_id,))
		return True
	except Exception:
		return False




def _normalizar_texto(texto: Any) -> str:
	"""Normaliza texto para comparar nombres de columnas de forma tolerante."""
	if texto is None:
		return ""

	valor = str(texto).strip().lower()
	valor = "".join(
		c for c in unicodedata.normalize("NFKD", valor) if not unicodedata.combining(c)
	)
	return " ".join(valor.split())


def _validar_archivo_excel() -> None:
	if not ARCHIVO_NOTAS.exists():
		raise FileNotFoundError(f"No se encontró el archivo Excel en: {ARCHIVO_NOTAS}")


@lru_cache(maxsize=1)
def _cargar_hojas() -> tuple[pd.DataFrame, pd.DataFrame]:
	"""Carga y cachea las hojas requeridas del archivo Notas.xlsx."""
	_validar_archivo_excel()

	puntajes_df = pd.read_excel(ARCHIVO_NOTAS, sheet_name=HOJA_PUNTAJES)
	indicadores_df = pd.read_excel(ARCHIVO_NOTAS, sheet_name=HOJA_INDICADORES)
	return puntajes_df, indicadores_df


def _buscar_columna(df: pd.DataFrame, candidatos: list[str], hoja: str) -> str:
	columnas_normalizadas = {_normalizar_texto(col): col for col in df.columns}

	for candidato in candidatos:
		col_real = columnas_normalizadas.get(_normalizar_texto(candidato))
		if col_real is not None:
			return col_real

	raise KeyError(
		f"No se encontró ninguna de las columnas {candidatos} en la hoja '{hoja}'. "
		f"Columnas disponibles: {list(df.columns)}"
	)


def _valor_numerico(valor: Any) -> float | None:
	if pd.isna(valor):
		return None

	try:
		return float(valor)
	except (TypeError, ValueError):
		return None


def obtener_alumnos() -> list[str]:
	"""Retorna los nombres únicos de alumnos presentes en la hoja Puntajes."""
	puntajes_df, _ = _cargar_hojas()

	col_alumno = _buscar_columna(
		puntajes_df,
		["Alumno", "Nombre", "Nombre alumno", "Estudiante"],
		HOJA_PUNTAJES,
	)

	alumnos = (
		puntajes_df[col_alumno]
		.dropna()
		.astype(str)
		.str.strip()
	)
	alumnos = alumnos[alumnos != ""]

	return sorted(alumnos.unique().tolist())


def obtener_brechas_alumno(nombre_alumno: str) -> dict[str, Any]:
	"""
	Obtiene las brechas del alumno y su resumen de puntajes.

	Retorna un diccionario con:
	- alumno
	- indicadores_brecha: lista de códigos con logro < 0.60
	- puntaje_total
	- nota_final
	"""
	if not nombre_alumno or not str(nombre_alumno).strip():
		raise ValueError("Debe proporcionar un nombre de alumno válido.")

	alumno_objetivo = str(nombre_alumno).strip()
	alumno_objetivo_norm = _normalizar_texto(alumno_objetivo)

	puntajes_df, indicadores_df = _cargar_hojas()

	col_alumno_puntajes = _buscar_columna(
		puntajes_df,
		["Alumno", "Nombre", "Nombre alumno", "Estudiante"],
		HOJA_PUNTAJES,
	)
	col_puntaje_total = _buscar_columna(
		puntajes_df,
		["Puntaje total", "Total", "Puntaje", "Puntaje final"],
		HOJA_PUNTAJES,
	)
	col_nota_final = _buscar_columna(
		puntajes_df,
		["Nota final", "Nota", "Calificacion final", "Calificación final"],
		HOJA_PUNTAJES,
	)

	col_alumno_indicadores = _buscar_columna(
		indicadores_df,
		["Alumno", "Nombre", "Nombre alumno", "Estudiante"],
		HOJA_INDICADORES,
	)
	col_codigo_indicador = _buscar_columna(
		indicadores_df,
		["Codigo", "Código", "Indicador", "Codigo indicador", "Código indicador"],
		HOJA_INDICADORES,
	)
	col_logro = _buscar_columna(
		indicadores_df,
		["Logro", "% Logro", "Porcentaje logro", "Valor logro"],
		HOJA_INDICADORES,
	)

	puntajes_filtrado = puntajes_df[
		puntajes_df[col_alumno_puntajes].astype(str).map(_normalizar_texto) == alumno_objetivo_norm
	]
	if puntajes_filtrado.empty:
		raise ValueError(f"No se encontró al alumno '{alumno_objetivo}' en la hoja '{HOJA_PUNTAJES}'.")

	fila_puntajes = puntajes_filtrado.iloc[0]
	puntaje_total = _valor_numerico(fila_puntajes[col_puntaje_total])
	nota_final = _valor_numerico(fila_puntajes[col_nota_final])

	indicadores_filtrado = indicadores_df[
		indicadores_df[col_alumno_indicadores].astype(str).map(_normalizar_texto)
		== alumno_objetivo_norm
	]

	indicadores_brecha: list[str] = []
	for _, fila in indicadores_filtrado.iterrows():
		logro = _valor_numerico(fila[col_logro])
		if logro is not None and logro < UMBRAL_APROBACION:
			codigo = str(fila[col_codigo_indicador]).strip()
			if codigo and codigo.lower() != "nan":
				indicadores_brecha.append(codigo)

	return {
		"alumno": alumno_objetivo,
		"indicadores_brecha": indicadores_brecha,
		"puntaje_total": puntaje_total,
		"nota_final": nota_final,
	}
