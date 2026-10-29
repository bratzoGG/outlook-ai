# Outlook AI — atajos y funciones

Referencia rápida. Se abre desde el acceso directo **«Atajos Outlook AI»** del escritorio.

## Atajos globales

Funcionan con Outlook clásico abierto y un correo seleccionado.

| Atajo | Función | Qué devuelve |
| --- | --- | --- |
| `Ctrl+Alt+1` | Redactar respuesta **formal** | Solo el cuerpo, sin asunto ni firma |
| `Ctrl+Alt+2` | **Resumir** | Frase de contexto + 2-5 viñetas + «Acciones» |
| `Ctrl+Alt+3` | **Traducir al español** | Solo la traducción, respetando formato y nombres |
| `Ctrl+Alt+4` | **Clasificar y priorizar** | Prioridad, Motivo, Requiere respuesta, Acción |
| `Ctrl+Alt+5` | Redactar respuesta **cercana** | Igual que el 1, con tuteo |
| `Ctrl+Alt+6` | Redactar respuesta **muy breve** | Dos o tres frases |
| `Ctrl+Alt+7` | **Traducir al inglés** | Solo la traducción |

Si seleccionas varios correos, envía hasta 10 (útil para resumir un hilo).

## Panel de resultado

| Botón | Qué hace |
| --- | --- |
| **Copiar** | Lleva el texto al portapapeles |
| **Crear borrador de respuesta** | Abre «Responder a todos» con el texto encima del correo citado. Solo en los atajos 1, 5 y 6 |
| **Cerrar** | Descarta sin tocar nada |

El texto es editable antes de usarlo y nada se envía solo.

## Menú Inicio → Outlook AI

| Elemento | Función |
| --- | --- |
| Configurar motor y clave | Elige motor (Ollama, Groq, Gemini, DeepSeek), modelo, dirección y clave |
| 1 … 7 | Los mismos atajos, por si uno no responde |

## Cambiar de motor (Ollama ↔ Gemini)

1. Inicio → **Outlook AI** → *Configurar motor y clave*.
2. Elige el motor en la lista. El modelo y la dirección se rellenan solos.
3. **Guardar**. No hace falta reiniciar: el siguiente atajo ya usa el motor nuevo.

Cada motor guarda **su propia clave**: si vuelves a Ollama y luego a Gemini, la clave de
Gemini sigue ahí y no hay que volver a pegarla. Ollama no necesita clave.

Recomendación de uso: **Ollama para correo de clientes** (nada sale del equipo) y **Gemini o
Groq solo para textos no sensibles**, porque en las capas gratuitas el proveedor puede usar
los datos para mejorar sus productos.

## Diagnóstico (PowerShell)

Primero define la ruta del script:

```powershell
$s = "$env:LOCALAPPDATA\OutlookAI\outlook-ai.ps1"
```

| Comando | Para qué |
| --- | --- |
| `powershell -File $s -SelfTest` | Revisa configuración, Outlook y conexión con el motor |
| `powershell -File $s -CheckKey` | Llamada mínima real: confirma que el modelo responde |
| `powershell -File $s -Mode summary -DryRun` | Muestra el prompt exacto sin llamar al modelo |
| `install.ps1 -Uninstall` | Quita los atajos y la carpeta de instalación |

## Estado actual

- **Motor**: Ollama local, modelo `qwen2.5:7b`. Coste por uso: **cero**.
- **Privacidad**: los correos no salen del equipo.
- **Velocidad**: 5-10 s por correo (el 7B); 2-4 s con `qwen2.5:3b`.
- **Memoria**: el 7B usa ~5 GB de RAM y los libera tras 5 minutos sin uso.
- **Ollama debe estar corriendo**: arranca con Windows (icono en la bandeja). Si falla la conexión, ábrelo desde Inicio.

## Problemas frecuentes

| Síntoma | Solución |
| --- | --- |
| «Falta configurar el motor» | Inicio → Outlook AI → Configurar motor y clave |
| Error de conexión con el motor | Abre Ollama desde Inicio; comprueba `powershell -File $s -CheckKey` |
| El atajo no responde | Cierra sesión y vuelve a entrar (Windows registra los atajos al iniciar sesión) |
| Un atajo deja de funcionar al instalar otra app | Otra aplicación usa el mismo `Ctrl+Alt+N`; cambia el suyo |

## Rutas

| Qué | Dónde |
| --- | --- |
| Script | `%LOCALAPPDATA%\OutlookAI\outlook-ai.ps1` |
| Configuración y clave cifrada | `%APPDATA%\outlook-ai\config.json` |
| Respaldo de la config DeepSeek | `%APPDATA%\outlook-ai\config.deepseek.json` |
| Modelos locales | `%USERPROFILE%\.ollama\models` |
| Proyecto (código fuente) | la carpeta donde descomprimiste el repositorio |
