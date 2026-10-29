# Changelog

## 1.0.0 — 2026-10-29

Primera versión pública.

- Asistente dentro de Outlook clásico vía automatización COM, sin add-in ni instalación de
  Office.
- Cuatro modos: redactar respuesta, resumir, traducir y clasificar, con tres tonos de
  redacción (formal, cercano, breve).
- **Cuatro motores intercambiables** desde la interfaz: Ollama (local, gratis), Groq y Google
  Gemini (capa gratuita) y DeepSeek (pago por uso). Todos mediante API compatible con OpenAI.
- Siete atajos globales en el menú Inicio (`Ctrl+Alt+1..7`) y panel de resultado editable con
  *Copiar* y *Crear borrador de respuesta*.
- Clave de API cifrada con DPAPI, ligada a la cuenta de Windows.
- Instalador y desinstalador sin dependencias, con limpieza de atajos obsoletos.
- Diagnóstico integrado: `-SelfTest`, `-CheckKey` y `-DryRun`.
