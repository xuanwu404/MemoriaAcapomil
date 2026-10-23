# Sistema de Retroalimentación de Instrumentos de Evaluación Asistido por IA

**Academia Politécnica Militar (ACAPOMIL) - Ejército de Chile**

Plataforma de prueba de concepto (PoC) diseñada para la automatización del razonamiento pedagógico y la formulación de planes de reforzamiento académico. El sistema emplea una arquitectura de Generación Aumentada por Recuperación (RAG) para procesar resultados de evaluaciones y emitir directrices técnicas fundamentadas en la doctrina militar e ingeniería de software indexada.

## ⚙️ Arquitectura del Sistema

El diseño lógico implementa un patrón de segregación de responsabilidades optimizado para cargas de trabajo de Inteligencia Artificial:

* **Capa de Presentación (Frontend):** Interfaz desarrollada en Streamlit (`frontend/app.py`). Gestiona el control de acceso, renderizado de telemetría académica, carga de doctrina y visualización estructurada de planes de trabajo.
* **Capa de Orquestación y LLM:** Construida sobre LangChain (`backend/orquestador.py`). Implementa un enrutamiento dual que soporta inferencia *On-Premise* aislada mediante Ollama (modelo `llama3.1`) para operaciones críticas sin conexión, o procesamiento en la nube vía API de OpenAI (`gpt-4o-mini`).
* **Motor de Embeddings:** Procesamiento vectorial local mediante la librería `sentence-transformers` en CPU. Utiliza el modelo `paraphrase-multilingual-MiniLM-L12-v2`, seleccionado por su alto rendimiento en similitud semántica para textos técnicos en español.
* **Persistencia Vectorial:** Almacén multidimensional ChromaDB (`data/vector_store/`). Encargado de la indexación y recuperación rápida de fragmentos de doctrina desde archivos PDF, DOCX y PPTX. Emplea métricas de similitud coseno con un umbral estricto (Tau = 0.60) para mitigar alucinaciones.
* **Persistencia Relacional:** Base de datos SQLite (`data/acapomil.db`). Estructura el modelo de dominio relacional: control de cuentas (RBAC), registro de unidades, matrícula de alumnos y matriz de trazabilidad de puntajes.

## 🖥️ Prerrequisitos de Sistema (Host)

Para asegurar la correcta compilación de librerías vectoriales y la ejecución del motor de inferencia local, la estación de trabajo debe contar con el siguiente software base:

1. **Python 3.10 o superior:** [Descargar instalador oficial](https://www.python.org/downloads/). *(Nota para Windows: Marcar obligatoriamente la casilla "Add Python to PATH" durante la instalación).*
2. **Git:** [Descargar cliente Git](https://git-scm.com/downloads) para el control de versiones y clonación del repositorio.
3. **Ollama:** [Descargar motor de inferencia](https://ollama.com/download) para habilitar la operación *On-Premise*.
4. **Modelo Llama 3.1:** Una vez instalado Ollama, abrir una terminal de sistema y ejecutar el siguiente comando para descargar los pesos del modelo neuronal:
   ```bash
   ollama pull llama3.1

Ejecutar los siguientes comandos en la terminal para inicializar el entorno aislado y desplegar la aplicación:

1. Clonar el repositorio

    ```Bash
    git clone https://github.com/xuanwu404/MemoriaAcapomil.git
    cd MemoriaAcapomil

2. Configurar el entorno virtual

    ```Bash
    python -m venv venv
    .\venv\Scripts\activate

3. Instalar dependencias
Instalar el ecosistema RAG y librerías de procesamiento de datos:

    ```Bash
    pip install -r requirements.txt

4. Configuración de Variables de Entorno
Crear el archivo .env en el directorio raíz para definir los parámetros de conexión LLM:

    ```Bash
    OPENAI_API_KEY=tu_api_key_aqui_si_usas_cloud
    LLM_PROVIDER=ollama
    OPENAI_MODEL=gpt-4o-mini

5. Ejecución del sistema
Inicializar el servidor local:

    ```Bash
    streamlit run frontend/app.py