<#
.SYNOPSIS
    Asistente de IA dentro de Outlook clásico: redacta, resume, traduce y clasifica.

.DESCRIPTION
    Lee el elemento seleccionado en Outlook (COM), construye un prompt según el modo y
    consulta un modelo. Soporta varios motores: DeepSeek (pago por uso), Groq y Google
    Gemini (capa gratuita) y Ollama (modelo local, sin coste ni cuenta).

.PARAMETER Mode
    reply | summary | translate | classify

.PARAMETER Tone
    formal | cercano | breve   (solo para reply)

.PARAMETER TargetLang
    Idioma destino para translate (por defecto español).

.PARAMETER Provider
    deepseek | groq | gemini | ollama | custom

.PARAMETER Setup
    Guarda la configuración: motor, modelo, dirección y clave.

.PARAMETER DryRun
    No usa Outlook ni la red: imprime el prompt que se enviaría.

.PARAMETER SelfTest
    Diagnóstico: configuración, Outlook por COM y conectividad con el motor.

.PARAMETER CheckKey
    Hace una llamada mínima real al motor y confirma que responde.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File outlook-ai.ps1 -Setup
    powershell -ExecutionPolicy Bypass -File outlook-ai.ps1 -Mode reply -Tone formal
#>
[CmdletBinding()]
param(
    [ValidateSet('reply', 'summary', 'translate', 'classify')]
    [string]$Mode = 'reply',

    [ValidateSet('formal', 'cercano', 'breve')]
    [string]$Tone = 'formal',

    [string]$TargetLang = 'español',

    [ValidateSet('deepseek', 'groq', 'gemini', 'ollama', 'custom')]
    [string]$Provider,

    [string]$Model,
    [string]$BaseUrl,

    [switch]$Setup,
    [switch]$DryRun,
    [switch]$SelfTest,
    [switch]$CheckKey,
    [string]$ConfigPath = (Join-Path $env:APPDATA 'outlook-ai\config.json')
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
try { Add-Type -AssemblyName System.Net.Http } catch { }

# ─────────────────────────────────────────────────────────────────────────────
# Motores disponibles
# ─────────────────────────────────────────────────────────────────────────────
function Get-Preset {
    param([string]$Name)
    switch ($Name) {
        'deepseek' { return [pscustomobject]@{ Id = 'deepseek'; Label = 'DeepSeek (pago por uso)'; BaseUrl = 'https://api.deepseek.com'; Model = 'deepseek-chat'; NeedsKey = $true } }
        'groq'     { return [pscustomobject]@{ Id = 'groq';     Label = 'Groq (capa gratuita)';    BaseUrl = 'https://api.groq.com/openai/v1'; Model = 'llama-3.3-70b-versatile'; NeedsKey = $true } }
        'gemini'   { return [pscustomobject]@{ Id = 'gemini';   Label = 'Google Gemini (capa gratuita)'; BaseUrl = 'https://generativelanguage.googleapis.com/v1beta/openai'; Model = 'gemini-3.5-flash'; NeedsKey = $true } }
        'ollama'   { return [pscustomobject]@{ Id = 'ollama';   Label = 'Ollama (local, gratis)';  BaseUrl = 'http://localhost:11434/v1'; Model = 'qwen2.5:7b'; NeedsKey = $false } }
        default    { return [pscustomobject]@{ Id = 'custom';   Label = 'Personalizado';           BaseUrl = ''; Model = ''; NeedsKey = $false } }
    }
}

function Get-ProviderList {
    return @('deepseek', 'groq', 'gemini', 'ollama')
}

# ─────────────────────────────────────────────────────────────────────────────
# Configuración
# ─────────────────────────────────────────────────────────────────────────────
function ConvertFrom-Dpapi {
    param([string]$Cipher)
    if ([string]::IsNullOrWhiteSpace($Cipher)) { return '' }
    try {
        $secure = ConvertTo-SecureString $Cipher
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } catch { return '' }
}

function ConvertTo-Dpapi {
    param([string]$Plain)
    if ([string]::IsNullOrWhiteSpace($Plain)) { return '' }
    $secure = ConvertTo-SecureString -String $Plain -AsPlainText -Force
    return ($secure | ConvertFrom-SecureString)
}

function Read-RawConfig {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try { return (Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json) } catch { return $null }
}

function Get-Config {
    param([string]$Path)
    $raw = Read-RawConfig -Path $Path
    if ($null -eq $raw) { return $null }
    try {
        $provider = 'deepseek'
        if ($raw.provider) { $provider = [string]$raw.provider }

        # Una clave por motor, para poder alternar sin volver a pegarlas.
        $keys = @{}
        if ($raw.PSObject.Properties['keys'] -and $null -ne $raw.keys) {
            foreach ($p in $raw.keys.PSObject.Properties) { $keys[$p.Name] = ConvertFrom-Dpapi ([string]$p.Value) }
        }
        # Compatibilidad con configuraciones antiguas (una sola clave).
        if ($raw.apiKey -and -not $keys.ContainsKey($provider)) { $keys[$provider] = ConvertFrom-Dpapi ([string]$raw.apiKey) }

        $preset = Get-Preset $provider
        $baseUrl = $preset.BaseUrl
        if ($raw.baseUrl) { $baseUrl = [string]$raw.baseUrl }
        $model = $preset.Model
        if ($raw.model) { $model = [string]$raw.model }
        $tone = 'formal'
        if ($raw.tone) { $tone = [string]$raw.tone }

        $currentKey = ''
        if ($keys.ContainsKey($provider)) { $currentKey = $keys[$provider] }

        return [pscustomobject]@{
            ApiKey   = $currentKey
            Keys     = $keys
            Provider = $provider
            BaseUrl  = $baseUrl
            Model    = $model
            Tone     = $tone
        }
    } catch {
        return $null
    }
}

function Save-Config {
    param([string]$Path, [string]$Provider, [string]$BaseUrl, [string]$Model, [string]$ApiKey, [string]$Tone)
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

    # Conservar las claves cifradas de los demás motores.
    $keys = @{}
    $raw = Read-RawConfig -Path $Path
    if ($null -ne $raw) {
        if ($raw.PSObject.Properties['keys'] -and $null -ne $raw.keys) {
            foreach ($p in $raw.keys.PSObject.Properties) { $keys[$p.Name] = [string]$p.Value }
        } elseif ($raw.apiKey -and $raw.provider) {
            $keys[[string]$raw.provider] = [string]$raw.apiKey
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ApiKey)) { $keys[$Provider] = ConvertTo-Dpapi $ApiKey }

    $encryptedCurrent = ''
    if ($keys.ContainsKey($Provider)) { $encryptedCurrent = $keys[$Provider] }

    $obj = [ordered]@{
        provider = $Provider
        baseUrl  = $BaseUrl
        model    = $Model
        tone     = $Tone
        apiKey   = $encryptedCurrent
        keys     = $keys
    }
    $json = $obj | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Show-Message {
    param([string]$Text, [string]$Title = 'Outlook AI', [string]$Icon = 'Information')
    [System.Windows.Forms.MessageBox]::Show($Text, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, $Icon) | Out-Null
}

# ─────────────────────────────────────────────────────────────────────────────
# Ventana de configuración
# ─────────────────────────────────────────────────────────────────────────────
if ($Setup) {
    $current = Get-Config -Path $ConfigPath
    $initProvider = 'deepseek'
    if ($current) { $initProvider = $current.Provider }
    $initPreset = Get-Preset $initProvider

    $win = New-Object System.Windows.Forms.Form
    $win.Text = 'Configuración de Outlook AI'
    $win.Size = New-Object System.Drawing.Size(560, 400)
    $win.StartPosition = 'CenterScreen'
    $win.FormBorderStyle = 'FixedDialog'
    $win.MaximizeBox = $false

    $y = 14

    $lblProv = New-Object System.Windows.Forms.Label
    $lblProv.Text = 'Motor:'
    $lblProv.Location = New-Object System.Drawing.Point(16, $y)
    $lblProv.Size = New-Object System.Drawing.Size(120, 20)
    $win.Controls.Add($lblProv)

    $cmbProv = New-Object System.Windows.Forms.ComboBox
    $cmbProv.Location = New-Object System.Drawing.Point(140, ($y - 2))
    $cmbProv.Size = New-Object System.Drawing.Size(380, 24)
    $cmbProv.DropDownStyle = 'DropDownList'
    $providerIds = Get-ProviderList
    foreach ($p in $providerIds) { [void]$cmbProv.Items.Add((Get-Preset $p).Label) }
    $initIndex = 0
    for ($i = 0; $i -lt $providerIds.Count; $i++) { if ($providerIds[$i] -eq $initProvider) { $initIndex = $i } }
    $cmbProv.SelectedIndex = $initIndex
    $win.Controls.Add($cmbProv)

    $y += 40
    $lblModel = New-Object System.Windows.Forms.Label
    $lblModel.Text = 'Modelo:'
    $lblModel.Location = New-Object System.Drawing.Point(16, $y)
    $lblModel.Size = New-Object System.Drawing.Size(120, 20)
    $win.Controls.Add($lblModel)

    $txtModel = New-Object System.Windows.Forms.TextBox
    $txtModel.Location = New-Object System.Drawing.Point(140, ($y - 2))
    $txtModel.Size = New-Object System.Drawing.Size(380, 24)
    if ($current) { $txtModel.Text = $current.Model } else { $txtModel.Text = $initPreset.Model }
    $win.Controls.Add($txtModel)

    $y += 40
    $lblUrl = New-Object System.Windows.Forms.Label
    $lblUrl.Text = 'Dirección:'
    $lblUrl.Location = New-Object System.Drawing.Point(16, $y)
    $lblUrl.Size = New-Object System.Drawing.Size(120, 20)
    $win.Controls.Add($lblUrl)

    $txtUrl = New-Object System.Windows.Forms.TextBox
    $txtUrl.Location = New-Object System.Drawing.Point(140, ($y - 2))
    $txtUrl.Size = New-Object System.Drawing.Size(380, 24)
    if ($current) { $txtUrl.Text = $current.BaseUrl } else { $txtUrl.Text = $initPreset.BaseUrl }
    $win.Controls.Add($txtUrl)

    $y += 40
    $lblKey = New-Object System.Windows.Forms.Label
    $lblKey.Text = 'Clave de API:'
    $lblKey.Location = New-Object System.Drawing.Point(16, $y)
    $lblKey.Size = New-Object System.Drawing.Size(120, 20)
    $win.Controls.Add($lblKey)

    $txtKey = New-Object System.Windows.Forms.TextBox
    $txtKey.Location = New-Object System.Drawing.Point(140, ($y - 2))
    $txtKey.Size = New-Object System.Drawing.Size(380, 24)
    $txtKey.UseSystemPasswordChar = $true
    if ($current -and $current.ApiKey) { $txtKey.Text = $current.ApiKey }
    $win.Controls.Add($txtKey)

    $y += 32
    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Location = New-Object System.Drawing.Point(140, $y)
    $lblHint.Size = New-Object System.Drawing.Size(380, 56)
    $lblHint.ForeColor = [System.Drawing.Color]::DimGray
    $win.Controls.Add($lblHint)

    $hints = @{
        deepseek = 'Pago por uso. Clave en platform.deepseek.com.'
        groq     = 'Gratis con límites. Clave en console.groq.com.'
        gemini   = 'Gratis con límites. Clave en aistudio.google.com.'
        ollama   = 'Local: sin clave ni internet. Instala Ollama y ejecuta: ollama pull qwen2.5:7b'
        custom   = 'Cualquier servidor compatible con la API de OpenAI.'
    }

    $providerIdsRef = $providerIds
    $storedKeys = @{}
    if ($current -and $current.Keys) { $storedKeys = $current.Keys }

    $applyPreset = {
        $id = $providerIdsRef[$cmbProv.SelectedIndex]
        $p = Get-Preset $id
        if ($p.Model) { $txtModel.Text = $p.Model }
        if ($p.BaseUrl) { $txtUrl.Text = $p.BaseUrl }
        $lblHint.Text = $hints[$id]
        $txtKey.Enabled = $p.NeedsKey
        # Recuperar la clave que ya estuviera guardada para ese motor.
        $txtKey.Text = ''
        if ($storedKeys.ContainsKey($id)) { $txtKey.Text = $storedKeys[$id] }
        if (-not $p.NeedsKey) { $txtKey.Clear() }
    }

    $cmbProv.Add_SelectedIndexChanged($applyPreset)
    & $applyPreset

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Guardar'
    $ok.Location = New-Object System.Drawing.Point(320, 320)
    $ok.Size = New-Object System.Drawing.Size(96, 30)
    $ok.DialogResult = 'OK'
    $win.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancelar'
    $cancel.Location = New-Object System.Drawing.Point(424, 320)
    $cancel.Size = New-Object System.Drawing.Size(96, 30)
    $cancel.DialogResult = 'Cancel'
    $win.Controls.Add($cancel)

    $win.AcceptButton = $ok
    $win.CancelButton = $cancel

    if ($win.ShowDialog() -eq 'OK') {
        $provId = $providerIdsRef[$cmbProv.SelectedIndex]
        $preset = Get-Preset $provId
        $key = ''
        if ($preset.NeedsKey) { $key = $txtKey.Text.Trim() }
        if ($preset.NeedsKey -and [string]::IsNullOrWhiteSpace($key)) {
            Show-Message 'Ese motor necesita una clave de API.' 'Outlook AI' 'Warning'
            exit 1
        }
        Save-Config -Path $ConfigPath -Provider $provId -BaseUrl $txtUrl.Text.Trim() `
            -Model $txtModel.Text.Trim() -ApiKey $key -Tone 'formal'
        Show-Message ("Configuración guardada.`n`nMotor: {0}`nModelo: {1}`n`nLas claves de los demás motores se conservan." -f $preset.Label, $txtModel.Text.Trim()) 'Outlook AI'
    }
    exit 0
}

# ─────────────────────────────────────────────────────────────────────────────
# Lectura del correo seleccionado
# ─────────────────────────────────────────────────────────────────────────────
function Get-SelectedMail {
    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    } catch {
        $outlook = New-Object -ComObject Outlook.Application
    }

    $explorer = $outlook.ActiveExplorer()
    if ($null -eq $explorer) { throw 'Abre la ventana principal de Outlook y selecciona un correo.' }

    $selection = $explorer.Selection
    if ($null -eq $selection -or $selection.Count -eq 0) {
        throw 'No hay ningún correo seleccionado. Selecciona uno en la lista y vuelve a intentarlo.'
    }

    $items = @()
    $max = [Math]::Min($selection.Count, 10)
    for ($i = 1; $i -le $max; $i++) {
        $item = $selection.Item($i)
        if ($item.Class -ne 43) { continue }   # 43 = olMail

        $fromAddr = ''
        try { $fromAddr = [string]$item.SenderEmailAddress } catch { $fromAddr = '' }

        $received = ''
        try { $received = ([datetime]$item.ReceivedTime).ToString('yyyy-MM-dd HH:mm') } catch { $received = '' }

        $items += [pscustomobject]@{
            Subject  = [string]$item.Subject
            From     = [string]$item.SenderName
            FromAddr = $fromAddr
            To       = [string]$item.To
            Received = $received
            Body     = [string]$item.Body
            Raw      = $item
        }
    }
    if ($items.Count -eq 0) { throw 'El elemento seleccionado no es un correo.' }
    return $items
}

# ─────────────────────────────────────────────────────────────────────────────
# Prompts
# ─────────────────────────────────────────────────────────────────────────────
function Format-MailBlob {
    param($Mails)
    $parts = @()
    foreach ($m in $Mails) {
        $body = $m.Body
        if ($body.Length -gt 12000) { $body = $body.Substring(0, 12000) + "`n[...texto recortado...]" }
        $parts += "De: $($m.From) <$($m.FromAddr)>`nPara: $($m.To)`nFecha: $($m.Received)`nAsunto: $($m.Subject)`n`n$body"
    }
    return ($parts -join "`n`n----------`n`n")
}

function Get-Prompt {
    param([string]$Mode, $Mails, [string]$Tone, [string]$TargetLang)

    $blob = Format-MailBlob -Mails $Mails
    switch ($Mode) {
        'reply' {
            $estilo = switch ($Tone) {
                'formal'  { 'profesional y cordial, en español neutro, sin florituras' }
                'cercano' { 'cercano y natural, tuteando, en español neutro' }
                'breve'   { 'muy breve: dos o tres frases como máximo' }
            }
            $system = "Eres un asistente que redacta respuestas de correo. Escribe en el idioma del correo original salvo que sea imposible. Devuelve SOLO el cuerpo de la respuesta, sin línea de asunto, sin «Asunto:», sin comentarios sobre lo que haces y sin firmar. Estilo: $estilo."
            $user = "Redacta una respuesta a este correo:`n`n$blob"
        }
        'summary' {
            $system = 'Resume correos con precisión. Devuelve un resumen breve en español con esta forma: una frase de qué trata, luego 2 a 5 viñetas con los puntos clave, y una sección «Acciones» con lo que se espera de mí (si no hay nada, escribe «Ninguna»). No inventes datos.'
            $user = "Resume este correo:`n`n$blob"
        }
        'translate' {
            $system = "Traduce al $TargetLang. Conserva el formato, los saltos de línea y los nombres propios. No añadas comentarios ni expliques nada: devuelve solo la traducción."
            $user = $blob
        }
        'classify' {
            $system = 'Clasificas correos. Devuelve exactamente estas cuatro líneas, en español: «Prioridad: Alta|Media|Baja», «Motivo: …», «Requiere respuesta: Sí|No», «Acción sugerida: …». Sé concreto y no inventes.'
            $user = "Clasifica este correo:`n`n$blob"
        }
    }
    return [pscustomobject]@{ System = $system; User = $user }
}

# ─────────────────────────────────────────────────────────────────────────────
# Llamada al modelo (API compatible con OpenAI: DeepSeek, Groq, Gemini, Ollama)
# ─────────────────────────────────────────────────────────────────────────────
function New-WorkingWindow {
    param([string]$Text = 'Consultando al modelo…')
    $f = New-Object System.Windows.Forms.Form
    $f.FormBorderStyle = 'FixedToolWindow'
    $f.StartPosition = 'CenterScreen'
    $f.Size = New-Object System.Drawing.Size(360, 110)
    $f.TopMost = $true
    $f.ShowInTaskbar = $false
    $f.Text = 'Outlook AI'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.Location = New-Object System.Drawing.Point(16, 16)
    $lbl.Size = New-Object System.Drawing.Size(320, 24)
    $f.Controls.Add($lbl)

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Style = 'Marquee'
    $bar.MarqueeAnimationSpeed = 30
    $bar.Location = New-Object System.Drawing.Point(16, 48)
    $bar.Size = New-Object System.Drawing.Size(320, 18)
    $f.Controls.Add($bar)

    $f.Show()
    [System.Windows.Forms.Application]::DoEvents()
    return $f
}

function Invoke-Model {
    param(
        [string]$BaseUrl,
        [string]$ApiKey,
        [string]$Model,
        [string]$System,
        [string]$User,
        [switch]$ShowProgress
    )

    $payload = @{
        model    = $Model
        messages = @(
            @{ role = 'system'; content = $System },
            @{ role = 'user'; content = $User }
        )
        temperature = 0.3
        stream      = $false
    } | ConvertTo-Json -Depth 8 -Compress

    $endpoint = $BaseUrl.TrimEnd('/') + '/chat/completions'

    $client = New-Object System.Net.Http.HttpClient
    $client.Timeout = [TimeSpan]::FromSeconds(180)
    if ($ApiKey) {
        $client.DefaultRequestHeaders.Authorization =
            New-Object System.Net.Http.Headers.AuthenticationHeaderValue('Bearer', $ApiKey)
    }
    $content = New-Object System.Net.Http.StringContent($payload, [System.Text.Encoding]::UTF8, 'application/json')

    $splash = $null
    if ($ShowProgress) { $splash = New-WorkingWindow }

    try {
        $task = $client.PostAsync($endpoint, $content)
        while (-not $task.IsCompleted) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 60
        }

        try { $response = $task.GetAwaiter().GetResult() }
        catch {
            $inner = $_.Exception
            while ($inner.InnerException) { $inner = $inner.InnerException }
            throw "No se pudo conectar con $endpoint.`n`n$($inner.Message)"
        }

        $text = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

        if (-not $response.IsSuccessStatusCode) {
            $code = [int]$response.StatusCode
            if ($code -eq 400) { throw "La petición fue rechazada (HTTP 400). Suele ser la clave de API incorrecta o el nombre del modelo mal escrito.`n`n$text" }
            if ($code -eq 401) { throw "Clave de API no válida (HTTP 401). Revisa la configuración." }
            if ($code -eq 403) { throw "Acceso denegado (HTTP 403). La clave puede no tener permiso para ese modelo.`n`n$text" }
            if ($code -eq 402) { throw "La cuenta no tiene saldo (HTTP 402). Recarga o cambia a un motor gratuito." }
            if ($code -eq 404) { throw "Modelo o dirección no encontrados (HTTP 404). Comprueba el modelo «$Model» y la dirección «$BaseUrl»." }
            if ($code -eq 429) { throw "Límite de peticiones alcanzado (HTTP 429). Espera un momento o cambia de motor." }
            throw "El servidor respondió HTTP $code.`n`n$text"
        }

        $json = $text | ConvertFrom-Json
        if (-not $json.choices -or $json.choices.Count -eq 0) { throw "Respuesta vacía del modelo.`n`n$text" }
        $answer = [string]$json.choices[0].message.content

        # Los modelos pequeños a veces pegan frases («cita.Gracias») o dejan saltos
        # de línea a mitad de párrafo. Se normaliza el espaciado sin tocar el contenido.
        $answer = [regex]::Replace($answer, '([.!?…])([A-ZÁÉÍÓÚÑÜ])', '$1 $2')
        $answer = [regex]::Replace($answer, '(,)([A-ZÁÉÍÓÚÑÜ])', '$1 $2')
        $answer = $answer -replace "`r`n", "`n"
        $answer = [regex]::Replace($answer, "`n{3,}", "`n`n")
        return $answer.Trim()
    } finally {
        if ($null -ne $splash) { $splash.Close(); $splash.Dispose() }
        $content.Dispose()
        $client.Dispose()
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Ventana de resultado
# ─────────────────────────────────────────────────────────────────────────────
function Show-Result {
    param([string]$Text, [string]$Title, $MailForReply)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(760, 560)
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $form.MinimumSize = New-Object System.Drawing.Size(520, 360)

    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline = $true
    $box.ScrollBars = 'Vertical'
    $box.Dock = 'Fill'
    $box.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $box.Text = $Text
    $box.Select(0, 0)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Bottom'
    $panel.Height = 56

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Copiar'
    $btnCopy.Location = New-Object System.Drawing.Point(12, 12)
    $btnCopy.Size = New-Object System.Drawing.Size(110, 32)
    $btnCopy.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($box.Text)
        $btnCopy.Text = 'Copiado ✓'
    })
    $panel.Controls.Add($btnCopy)

    if ($null -ne $MailForReply) {
        $btnReply = New-Object System.Windows.Forms.Button
        $btnReply.Text = 'Crear borrador de respuesta'
        $btnReply.Location = New-Object System.Drawing.Point(132, 12)
        $btnReply.Size = New-Object System.Drawing.Size(220, 32)
        $btnReply.Add_Click({
            try {
                $draft = $MailForReply.ReplyAll()
                $escaped = [System.Security.SecurityElement]::Escape($box.Text)
                $draft.HTMLBody = '<div style="font-family:Segoe UI,sans-serif;font-size:11pt">' +
                    ($escaped -replace "`r?`n", '<br>') + '</div><br>' + $draft.HTMLBody
                $draft.Display()
                $form.Close()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("No se pudo crear el borrador: $($_.Exception.Message)", 'Outlook AI')
            }
        })
        $panel.Controls.Add($btnReply)
    }

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Cerrar'
    $btnClose.Location = New-Object System.Drawing.Point(640, 12)
    $btnClose.Size = New-Object System.Drawing.Size(96, 32)
    $btnClose.Anchor = 'Top,Right'
    $btnClose.Add_Click({ $form.Close() })
    $panel.Controls.Add($btnClose)

    $form.Controls.Add($box)
    $form.Controls.Add($panel)
    $form.Add_Shown({ $box.Focus() })
    [void]$form.ShowDialog()
}

# ─────────────────────────────────────────────────────────────────────────────
# Flujo principal
# ─────────────────────────────────────────────────────────────────────────────
$titles = @{
    reply     = 'Respuesta propuesta'
    summary   = 'Resumen del correo'
    translate = "Traducción a $TargetLang"
    classify  = 'Clasificación del correo'
}

try {
    $config = Get-Config -Path $ConfigPath

    # ── comprobación del motor ──
    if ($CheckKey) {
        if ($null -eq $config) {
            Write-Host 'Sin configuración. Ejecuta «Configurar clave de API» desde el menú Inicio.'
            exit 1
        }
        $preset = Get-Preset $config.Provider
        Write-Host ("Motor: {0}  Modelo: {1}" -f $preset.Label, $config.Model)
        Write-Host ("Dirección: {0}" -f $config.BaseUrl)
        try {
            $answer = Invoke-Model -BaseUrl $config.BaseUrl -ApiKey $config.ApiKey -Model $config.Model `
                -System 'Responde únicamente con la palabra OK.' -User 'ping'
            $clean = ($answer -replace "`r?`n", ' ').Trim()
            Write-Host ("Correcto. Respuesta del modelo: {0}" -f $clean)
            exit 0
        } catch {
            Write-Host ("Fallo: {0}" -f $_.Exception.Message)
            exit 1
        }
    }

    # ── diagnóstico ──
    if ($SelfTest) {
        Write-Host '== Outlook AI: diagnóstico =='
        if ($null -eq $config) {
            Write-Host 'Configuración: SIN CONFIGURAR'
        } else {
            $preset = Get-Preset $config.Provider
            Write-Host ("Configuración: {0}, modelo {1}" -f $preset.Label, $config.Model)
            Write-Host ("Dirección: {0}" -f $config.BaseUrl)
            if ($config.ApiKey) { Write-Host 'Clave: presente' } else { Write-Host 'Clave: no usada (motor local)' }
        }

        Write-Host 'Conexión a Outlook...'
        try {
            $o = $null
            try { $o = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
            catch { $o = New-Object -ComObject Outlook.Application }
            $exp = $o.ActiveExplorer()
            if ($null -eq $exp) {
                Write-Host '  Outlook responde, pero no hay ventana principal abierta.'
            } else {
                Write-Host ("  Outlook responde. Correos seleccionados: {0}" -f $exp.Selection.Count)
            }
        } catch {
            Write-Host ("  Outlook no disponible: {0}" -f $_.Exception.Message)
        }

        if ($null -ne $config) {
            Write-Host ("Conectividad con {0}..." -f $config.BaseUrl)
            try {
                $null = Invoke-Model -BaseUrl $config.BaseUrl -ApiKey $config.ApiKey -Model $config.Model `
                    -System 'Responde OK.' -User 'ping'
                Write-Host '  El modelo respondió correctamente.'
            } catch {
                Write-Host ("  {0}" -f $_.Exception.Message)
            }
        }
        Write-Host '== fin =='
        exit 0
    }

    # ── prompt en seco ──
    if ($DryRun) {
        $sample = @([pscustomobject]@{
            Subject = 'Reunión de presupuesto el jueves'; From = 'Ana Ruiz'; FromAddr = 'ana@example.com'
            To = 'Equipo de compras'; Received = '2026-10-28 09:15'
            Body = 'Hola, ¿puedes confirmar si llegamos al jueves con la propuesta? Necesito el número final para la reunión con el banco. Gracias.'
        })
        $p = Get-Prompt -Mode $Mode -Mails $sample -Tone $Tone -TargetLang $TargetLang
        Write-Host '=== SYSTEM ==='
        Write-Host $p.System
        Write-Host ''
        Write-Host '=== USER ==='
        Write-Host $p.User
        exit 0
    }

    # ── flujo normal ──
    if ($null -eq $config) {
        Show-Message ("Falta configurar el motor.`n`n" +
            "Abre el menú Inicio, busca «Outlook AI» y ejecuta «Configurar clave de API».`n`n" +
            "Puedes usar un modelo local gratuito (Ollama) o una clave con capa gratuita.") 'Outlook AI' 'Warning'
        exit 1
    }

    $effectiveProvider = $config.Provider
    if ($PSBoundParameters.ContainsKey('Provider')) { $effectiveProvider = $Provider }
    $effectivePreset = Get-Preset $effectiveProvider

    $effectiveBaseUrl = $config.BaseUrl
    if ($PSBoundParameters.ContainsKey('BaseUrl')) { $effectiveBaseUrl = $BaseUrl }
    elseif ($effectiveProvider -ne $config.Provider) { $effectiveBaseUrl = $effectivePreset.BaseUrl }

    $effectiveModel = $config.Model
    if ($PSBoundParameters.ContainsKey('Model')) { $effectiveModel = $Model }
    elseif ($effectiveProvider -ne $config.Provider) { $effectiveModel = $effectivePreset.Model }

    $mails = Get-SelectedMail
    $prompt = Get-Prompt -Mode $Mode -Mails $mails -Tone $Tone -TargetLang $TargetLang
    $answer = Invoke-Model -BaseUrl $effectiveBaseUrl -ApiKey $config.ApiKey -Model $effectiveModel `
        -System $prompt.System -User $prompt.User -ShowProgress

    $replyTo = if ($Mode -eq 'reply') { $mails[0].Raw } else { $null }
    Show-Result -Text $answer -Title $titles[$Mode] -MailForReply $replyTo
} catch {
    Show-Message $_.Exception.Message 'Outlook AI' 'Error'
    exit 1
}
