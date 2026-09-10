from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from typing import Iterable
import math
import os

# Bloqueo de telemetría de ChromaDB (Evita errores de capture() en consola)
os.environ["ANONYMIZED_TELEMETRY"] = "False"

import docx2txt
# Importación estándar y moderna:
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_core.documents import Document
from langchain_core.embeddings import Embeddings
from langchain_community.vectorstores import Chroma

try:
	from langchain_openai import OpenAIEmbeddings
except ImportError:  # pragma: no cover - dependencia opcional
	OpenAIEmbeddings = None


ROOT_DIR = Path(__file__).resolve().parent.parent
DOCTRINA_DIR = ROOT_DIR / "data" / "doctrina"
VECTOR_STORE_DIR = ROOT_DIR / "data" / "vector_store"
COLLECTION_NAME = "acapomil_doctrina"
TAU_SIMILITUD = 0.75


class HashEmbeddings(Embeddings):
	"""Embeddings deterministas para desarrollo local sin servicios externos."""

	def __init__(self, dim: int = 256) -> None:
		self.dim = dim

	def _embed_text(self, text: str) -> list[float]:
		vector = [0.0] * self.dim
		if not text:
			return vector

		tokens = text.lower().split()
		for token in tokens:
			digest = sha256(token.encode("utf-8")).digest()
			idx = int.from_bytes(digest[:4], "big") % self.dim
			sign = 1.0 if digest[4] % 2 == 0 else -1.0
			vector[idx] += sign

		norm = math.sqrt(sum(v * v for v in vector))
		if norm == 0:
			return vector
		return [v / norm for v in vector]

	def embed_documents(self, texts: list[str]) -> list[list[float]]:
		return [self._embed_text(text) for text in texts]

	def embed_query(self, text: str) -> list[float]:
		return self._embed_text(text)


@dataclass(frozen=True)
class ResultadoRecuperacion:
	fragmento: str
	similitud_coseno: float
	metadata: dict


def get_embeddings() -> Embeddings:
	"""
	Entrega el modelo de embeddings.

	Prioriza OpenAI cuando hay API key; en caso contrario usa embeddings locales
	deterministas para mantener el pipeline funcional sin dependencias externas.
	"""
	openai_api_key = os.getenv("OPENAI_API_KEY")
	if openai_api_key and OpenAIEmbeddings is not None:
		return OpenAIEmbeddings(model="text-embedding-3-small", api_key=openai_api_key)

	return HashEmbeddings()


def _extraer_texto_pdf(path_archivo: Path) -> str:
	from pypdf import PdfReader

	try:
		reader = PdfReader(path_archivo)
		texto = ""
		for pagina in reader.pages:
			texto += (pagina.extract_text() or "") + "\n"
		return texto.strip()
	except Exception as e:
		print(f"Error al leer PDF {path_archivo}: {e}")
		return ""


def _extraer_texto_pptx(path_archivo: Path) -> str:
	from pptx import Presentation

	try:
		prs = Presentation(path_archivo)
		texto = ""
		for slide in prs.slides:
			for shape in slide.shapes:
				if hasattr(shape, "text") and shape.text.strip():
					texto += shape.text.strip() + "\n"
			if slide.notes_slide and slide.notes_slide.notes_text_frame:
				texto += slide.notes_slide.notes_text_frame.text.strip() + "\n"
		return texto.strip()
	except Exception as e:
		print(f"Error al leer PPTX {path_archivo}: {e}")
		return ""


def cargar_documentos_docx(origen: Path = DOCTRINA_DIR) -> list[Document]:
	"""Carga doctrina multiproposito desde DOCX, PDF y PowerPoint."""
	if not origen.exists():
		raise FileNotFoundError(f"No existe la carpeta de doctrina: {origen}")

	documentos: list[Document] = []
	extensiones_permitidas = {".docx", ".pdf", ".pptx", ".ppt"}
	for ruta in sorted(p for p in origen.iterdir() if p.is_file() and p.suffix.lower() in extensiones_permitidas):
		texto = ""
		sufijo = ruta.suffix.lower()
		if sufijo == ".docx":
			texto = docx2txt.process(str(ruta))
		elif sufijo == ".pdf":
			texto = _extraer_texto_pdf(ruta)
		elif sufijo in {".pptx", ".ppt"}:
			texto = _extraer_texto_pptx(ruta)

		if not texto or not texto.strip():
			continue

		documentos.append(
			Document(
				page_content=texto,
				metadata={
					"source": str(ruta),
					"filename": ruta.name,
				},
			)
		)

	if not documentos:
		raise ValueError(
			f"No se encontraron documentos DOCX/PDF/PPTX/PPT con contenido en {origen}"
		)

	return documentos


def crear_segmentador() -> RecursiveCharacterTextSplitter:
	"""Configura el segmentador solicitado para chunking de doctrina."""
	return RecursiveCharacterTextSplitter(
		chunk_size=500,
		chunk_overlap=50,
		separators=["\n\n", "\n", ". ", " ", ""],
	)


def segmentar_documentos(documentos: Iterable[Document]) -> list[Document]:
	splitter = crear_segmentador()
	return splitter.split_documents(list(documentos))


def construir_vector_store(
	documentos: list[Document],
	persist_directory: Path = VECTOR_STORE_DIR,
	collection_name: str = COLLECTION_NAME,
) -> Chroma:
	"""Crea y persiste Chroma en data/vector_store a partir de los documentos."""
	persist_directory.mkdir(parents=True, exist_ok=True)

	embeddings = get_embeddings()
	chunks = segmentar_documentos(documentos)

	vector_store = Chroma.from_documents(
		documents=chunks,
		embedding=embeddings,
		collection_name=collection_name,
		persist_directory=str(persist_directory),
	)
	return vector_store


def cargar_vector_store(
	persist_directory: Path = VECTOR_STORE_DIR,
	collection_name: str = COLLECTION_NAME,
) -> Chroma:
	"""Carga una colección existente de Chroma desde data/vector_store."""
	if not persist_directory.exists():
		raise FileNotFoundError(
			f"No existe la carpeta del vector store: {persist_directory}. "
			"Ejecute primero la ingestión."
		)

	return Chroma(
		collection_name=collection_name,
		embedding_function=get_embeddings(),
		persist_directory=str(persist_directory),
	)


def _similitud_coseno(vec_a: list[float], vec_b: list[float]) -> float:
	if not vec_a or not vec_b or len(vec_a) != len(vec_b):
		return 0.0

	producto = sum(a * b for a, b in zip(vec_a, vec_b))
	norma_a = math.sqrt(sum(a * a for a in vec_a))
	norma_b = math.sqrt(sum(b * b for b in vec_b))

	if norma_a == 0 or norma_b == 0:
		return 0.0

	return producto / (norma_a * norma_b)


def recuperar_fragmentos_relevantes(
	pregunta: str,
	tau: float = TAU_SIMILITUD,
	k: int = 4,
	vector_store: Chroma | None = None,
) -> list[ResultadoRecuperacion]:
	"""
	Recupera fragmentos y valida similitud coseno >= tau antes de retornarlos.
	"""
	if not pregunta or not pregunta.strip():
		raise ValueError("La pregunta no puede estar vacía.")

	store = vector_store or cargar_vector_store()
	embeddings = get_embeddings()

	candidatos = store.similarity_search(query=pregunta, k=k)
	if not candidatos:
		return []

	emb_pregunta = embeddings.embed_query(pregunta)
	resultados: list[ResultadoRecuperacion] = []

	for doc in candidatos:
		emb_doc = embeddings.embed_query(doc.page_content)
		similitud = _similitud_coseno(emb_pregunta, emb_doc)
		if similitud >= tau:
			resultados.append(
				ResultadoRecuperacion(
					fragmento=doc.page_content,
					similitud_coseno=similitud,
					metadata=doc.metadata,
				)
			)

	resultados.sort(key=lambda item: item.similitud_coseno, reverse=True)
	return resultados


def ingest_arreglo_doctrina() -> Chroma:
	"""Ingesta completa: carga doctrina multiproposito y persiste la base vectorial."""
	documentos = cargar_documentos_docx()
	return construir_vector_store(documentos=documentos)
