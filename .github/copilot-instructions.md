# Contexto del Proyecto: Sistema de Retroalimentación RAG ACAPOMIL
- **Usuario**: Xuanwu (Ingeniero Politécnico Militar).
- **Objetivo**: MVP Web para automatizar retroalimentación pedagógica basada en competencias.
- **Autoridad Pedagógica**: Basado en el manual MAE-01006 e indicadores del profesor MAY (IPM) Pablo Varas Briones.
- **Reglas Críticas de Negocio**:
  - Estándar de aprobación mínimo por ítem/indicador: 60% (0.60).
  - Escala de notas: 1.0 a 7.0 (60 puntos = Nota 4.0).
  - Mitigación de Alucinaciones: Controlar mediante similitud de coseno con un umbral tau >= 0.75.
  - Despliegue Dual: Soportar conmutación entre OpenAI API (o Claude) y Ollama (Llama3/Mistral local).
- **Formatos de Entrada**: Parsear 'Notas.xlsx' (hojas: 'Puntajes' e 'Indicadores de desempeño') para extraer brechas automáticas por alumno.