# Sistema de Retroalimentación de Instrumentos de Evaluación Asistido por IA

**Academia Politécnica Militar (ACAPOMIL) - Ejército de Chile**

Diseño preliminar de un sistema RAG (Generación Aumentada por Recuperación) para la automatización del análisis pedagógico y la formulación de planes de refuerzo técnico.

## ⚙️ Arquitectura del Sistema
El sistema opera bajo un patrón modular segregado para garantizar el aislamiento de procesos y la soberanía de los datos:
* **Presentación (Frontend):** Interfaz desarrollada en Streamlit (`frontend/app.py`) para el renderizado de telemetría académica, carga de doctrina y visualización de planes de trabajo.
* **Orquestación y LLM:** Lógica de control estructurada sobre LangChain. Soporta inferencia local aisalda mediante Ollama (modelo `llama3.1`) o procesamiento externo vía API de OpenAI.
* **Motor de Embeddings:** Procesamiento vectorial local mediante `sentence-transformers` en CPU, empleando el modelo `paraphrase-multilingual-MiniLM-L12-v2` optimizado para la semántica en español[cite: 240].
* **Persistencia Vectorial:** Almacén ChromaDB configurado para la indexación y recuperación de doctrina militar desde archivos PDF, DOCX y PPTX, con particionamiento lógico por unidad académica[cite: 240].
* **Persistencia Relacional:** Base de datos SQLite (`acapomil.db`) encargada del control de cuentas, unidades, matrícula de alumnos y matriz de trazabilidad de puntajes.

## 🛠️ Requisitos de Entorno
* Python 3.10 o superior.
* Entorno virtual (`venv`).
* [Ollama](https://ollama.ai/) instalado y modelo `llama3.1` descargado localmente (para ejecución *On-Premise*).
* Git para control de versiones.

## 🚀 Guía de Despliegue Operacional

**1. Clonar el repositorio**
```bash
git clone [https://github.com/xuanwu404/MemoriaAcapomil.git](https://github.com/xuanwu404/MemoriaAcapomil.git)
cd MemoriaAcapomil