# Contribuir

Gracias por querer mejorar esto. Es un proyecto pequeño a propósito: **un solo archivo de
PowerShell** y un instalador.

## Antes de enviar un cambio

Comprueba que el script sigue siendo válido:

```powershell
powershell -NoProfile -Command "$e=$null; [void][System.Management.Automation.Language.Parser]::ParseFile('.\outlook-ai.ps1',[ref]$null,[ref]$e); if($e){$e; exit 1} else {'OK'}"
```

Y si el cambio toca la lógica de red o de prompts, pruébalo con un correo real antes del pull
request:

```powershell
powershell -File .\outlook-ai.ps1 -Mode summary -DryRun   # revisa el prompt
powershell -File .\outlook-ai.ps1 -CheckKey               # revisa el motor
```

## Ideas de contribución

- **Más idiomas de destino** para el modo traducir (`-TargetLang` ya es un parámetro).
- **Tonos nuevos**: hoy hay formal, cercano y breve; añadir «directivo» o «empático» es una
  rama más en `Get-Prompt`.
- **Lectura de adjuntos**: Outlook expone `Attachments`; requeriría extraer texto de PDF y
  DOCX, así que conviene plantearlo como opcional.
- **Otros motores**: cualquiera con API compatible con OpenAI funciona cambiando
  `BaseUrl` y `Model`; se puede añadir un preset más a `Get-Preset`.
- **Streaming**: hoy espera a la respuesta completa; mostrar el texto mientras se genera
  mejoraría la sensación de velocidad en modelos locales.

## Estilo

- PowerShell 5.1 compatible (nada de operadores de PowerShell 7).
- Los archivos `.ps1` en **UTF-8 con BOM**: sin BOM, Windows PowerShell 5.1 lee los acentos
  como ANSI y los textos salen corruptos.
- Comentarios y textos de interfaz en español.
- Nada de dependencias externas: el valor del proyecto es que funciona con lo que trae
  Windows.

## Reportar un problema

Incluye la salida de `-SelfTest` y qué motor usas. Si es un fallo de calidad del texto, pega
el prompt (`-DryRun`) y el resultado obtenido.
