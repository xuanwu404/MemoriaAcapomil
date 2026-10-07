# Sistema de Retroalimentación de Instrumentos de Evaluación Asistido por IA

**Academia Politécnica Militar (ACAPOMIL) - Ejército de Chile**

Plataforma de prueba de concepto (PoC) diseñada para la automatización del razonamiento pedagógico y la formulación de planes de reforzamiento académico. El sistema emplea una arquitectura de Generación Aumentada por Recuperación (RAG) para procesar resultados de evaluaciones y emitir directrices técnicas basadas estrictamente en la doctrina militar e ingeniería de software cargada en el sistema.

## ⚙️ Arquitectura del Sistema

El diseño lógico implementa un patrón de segregación de responsabilidades optimizado para cargas de trabajo de Inteligencia Artificial:

* **Capa de Presentación (Frontend):** Interfaz desarrollada en Streamlit (`frontend/app.py`). Gestiona el control de acceso, renderizado de telemetría académica, carga de doctrina y visualización estructurada de planes de trabajo.
* **Capa de Orquestación y LLM:** Construida sobre LangChain (`backend/orquestador.py`). Implementa un enrutamiento dual que soporta inferencia *On-Premise* aislada mediante Ollama (modelo `llama3.1`) para operaciones críticas sin conexión, o procesamiento en la nube vía API de OpenAI (`gpt-4o-mini`) para escalabilidad[cite: 239].
* **Motor de Embeddings:** Procesamiento vectorial local mediante la librería `sentence-transformers` en CPU. Utiliza el modelo `paraphrase-multilingual-MiniLM-L12-v2`, seleccionado por su alto rendimiento en similitud semántica para textos técnicos en español[cite: 240].
* **Persistencia Vectorial:** Almacén multidimensional ChromaDB (`data/vector_store/`). Encargado de la indexación y recuperación rápida de fragmentos de doctrina desde archivos PDF, DOCX y PPTX. Emplea métricas de similitud coseno con un umbral estricto (Tau = 0.60) para mitigar alucinaciones[cite: 240].
* **Persistencia Relacional:** Base de datos SQLite (`data/acapomil.db`). Estructura el modelo de dominio relacional: control de cuentas (RBAC), registro de unidades, matrícula de alumnos y matriz de trazabilidad de puntajes.

## 🛠️ Requisitos del Entorno

* Python 3.10 o superior.
* Entorno virtual (`venv`).
* [Ollama](https://ollama.ai/) instalado y en ejecución local con el modelo `llama3.1` (requerido para inferencia *On-Premise*).
* Git.

## 🚀 Guía de Despliegue Operacional

**1. Clonar el repositorio**
```bash
git clone [https://github.com/xuanwu404/MemoriaAcapomil.git](https://github.com/xuanwu404/MemoriaAcapomil.git)
cd MemoriaAcapomil