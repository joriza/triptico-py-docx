<#
.SYNOPSIS
    Genera los PDFs del kit "Tríptico para un emprendimiento casero de gelatina"
    a partir de los .md, de forma DETERMINÍSTICA.

.DESCRIPTION
    Pipeline:
      1. pandoc    : <doc>.md -> _<doc>.html (con custom-markdown-v2.css)
      2. wkhtmltopdf: _<doc>.html -> <doc>.pdf
      3. normalizar: reemplaza la fecha de creación del PDF por una constante

    POR QUÉ ES NECESARIO EL PASO 3
    wkhtmltopdf escribe la fecha y hora de creación en los metadatos del PDF.
    Dos ejecuciones con exactamente la misma entrada producen entonces archivos
    DISTINTOS (verificado: hashes distintos). El paso 3 reemplaza esa fecha por
    una constante de la MISMA longitud (23 caracteres), lo que preserva todos
    los offsets internos del PDF y hace que dos ejecuciones den archivos
    byte a byte idénticos.

    No cambiar la fecha por una de distinta longitud: desplazaría los offsets
    de la tabla xref y corrompería el archivo (comprobado: un reemplazo más
    corto dejó el PDF en 3 páginas en vez de 4).

    NOTAS DE MAQUETADO (no obvias, medidas sobre el PDF resultante)
    - wkhtmltopdf tiene "smart shrinking" activado por defecto y escala las
      longitudes del CSS por ~0.794 en el PDF. Los márgenes que recibe por
      flag, en cambio, NO se escalan. Por eso el margen lateral vive acá (ver
      $Margenes) y no en el CSS: el número de este script es el número que se
      imprime. Si el lateral se moviera al CSS, habría que multiplicarlo por
      1/0.794 para saber qué sale impreso.
    - Todos los márgenes de página vienen de los flags de $Margenes, 1:1. El
      CSS no define márgenes a propósito: se probó que el margin de body no
      sirve (el horizontal se aplica escalado por ese 0.794, así que el número
      declarado no es el impreso; el vertical se ignora por completo). El pie
      [page]/[topage] se dibuja dentro de los 10mm del flag inferior.
    - El CSS se enlaza con ruta relativa: el HTML temporal DEBE quedar en esta
      misma carpeta, o el estilo no carga y el PDF sale sin maquetar.

.PARAMETER Doc
    Nombres de documento sin la extensión .md. Si se omite, se generan todos.

.PARAMETER MargenLateral
    Margen lateral del PDF. Acepta cualquier unidad que entienda wkhtmltopdf
    (2.5cm, 25mm, 1in, 12px). Sale 1:1 en el papel: el valor que se pasa es el
    valor que se imprime. Por defecto: 2.5cm.

    Es el único control real del largo de línea y de las hojas gastadas. En A4 a
    una columna, las hojas crecen proporcionalmente a 1/ancho de columna: no hay
    mejora de legibilidad gratis salvo el primer tramo. Medido sobre este kit
    (hojas por documento):
        margen   cuaderno  guía  modelo  rúbrica   línea más larga
        1,42cm      8       8     4       4       130 caracteres
        2,00cm      8       -     -       -       122
        2,50cm      8       9     4       4       116
        3,00cm      9       -     -       -       111
    Referencia de legibilidad: hasta 90 caracteres por línea es cómodo, ~66 es lo
    ideal. En A4 a una columna con este motor no se baja de 90 sin gastar bastante
    más papel, y multicolumna NO es una opción: wkhtmltopdf (QtWebKit 4.8) la
    ignora, incluso con un contenedor explícito.

.PARAMETER Verify
    Genera cada PDF dos veces y compara los hashes SHA256 para comprobar el
    determinismo. Tarda el doble.

.EXAMPLE
    .\generar-pdf.ps1
    .\generar-pdf.ps1 -Doc rubrica-triptico-gelatina
    .\generar-pdf.ps1 -MargenLateral 1.42cm
    .\generar-pdf.ps1 -Verify
#>

[CmdletBinding()]
param(
    [string[]]$Doc,
    [string]$MargenLateral = '2.5cm',
    [switch]$Verify
)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

if ($MargenLateral -notmatch '^\d+(\.\d+)?(cm|mm|in|pt|px)$') {
    throw "MargenLateral '$MargenLateral' no tiene un formato válido. Ejemplos: 2.5cm, 25mm, 1in."
}

$CSS      = 'custom-markdown-v2.css'
$TodosDoc = @(
    'guia-docente-triptico-gelatina',
    'cuaderno-del-alumno-triptico-gelatina',
    'modelo-triptico-gelatina',
    'rubrica-triptico-gelatina'
)
$Docs = if ($PSBoundParameters.ContainsKey('Doc')) { $Doc } else { $TodosDoc }

# Margenes laterales: el valor de $MargenLateral, que sale 1:1 en el papel (ver
# la nota de smart shrinking del encabezado). El criterio para elegirlo y las
# mediciones que lo respaldan estan en la ayuda de -MargenLateral.
# Vertical: no se toca. 10mm de flag, que es donde se dibuja el pie de pagina.
$Margenes = @(
    '--margin-top',    '10mm',
    '--margin-bottom', '10mm',
    '--margin-left',   $MargenLateral,
    '--margin-right',  $MargenLateral
)

# Fecha fija, 23 caracteres: D:YYYYMMDDHHmmSS + HH'MM'
$FechaFija    = "D:20260101000000+00'00'"
$PatronFecha  = "D:\d{14}[+-]\d{2}'\d{2}'"

foreach ($herramienta in 'pandoc', 'wkhtmltopdf') {
    if (-not (Get-Command $herramienta -ErrorAction SilentlyContinue)) {
        throw "Falta '$herramienta' en el PATH."
    }
}

function Set-FechaDeterministica {
    param([Parameter(Mandatory)][string]$Ruta)

    # Latin1 (ISO-8859-1) mapea 1:1 cada byte, asi que sirve para tratar el PDF
    # como texto sin romper los bytes binarios.
    $enc    = [System.Text.Encoding]::GetEncoding(28591)
    $bytes  = [System.IO.File]::ReadAllBytes($Ruta)
    $texto  = $enc.GetString($bytes)
    $todos  = [regex]::Matches($texto, $PatronFecha)

    if ($todos.Count -ne 1) {
        throw "$Ruta`: se esperaba 1 fecha de creacion, se encontraron $($todos.Count). No se normaliza."
    }
    $encontrada = $todos[0].Value
    if ($encontrada.Length -ne $FechaFija.Length) {
        throw "$Ruta`: la fecha '$encontrada' mide $($encontrada.Length) y la fija mide $($FechaFija.Length). Igualar longitudes antes de continuar: un reemplazo de otra longitud desplaza los offsets del PDF."
    }

    $m       = $todos[0]
    $nuevo   = $texto.Substring(0, $m.Index) + $FechaFija + $texto.Substring($m.Index + $encontrada.Length)
    [System.IO.File]::WriteAllBytes($Ruta, $enc.GetBytes($nuevo))
}

function Invoke-GenerarPdf {
    param(
        [Parameter(Mandatory)][string]$Nombre,
        [Parameter(Mandatory)][string]$Salida
    )

    if (-not (Test-Path "$Nombre.md")) { throw "No existe $Nombre.md" }

    # El HTML temporal va en ESTA carpeta para que el CSS relativo resuelva.
    $tmpHtml = "_$Nombre.tmp.html"
    try {
        pandoc "$Nombre.md" `
            --from=markdown+task_lists+hard_line_breaks `
            --standalone `
            --css=$CSS `
            --variable="pagetitle:$Nombre" `
            -o $tmpHtml

        wkhtmltopdf --enable-local-file-access --quiet `
            @Margenes `
            --footer-center '[page]/[topage]' --footer-font-size 8 `
            $tmpHtml $Salida

        Set-FechaDeterministica -Ruta $Salida
    }
    finally {
        if (Test-Path $tmpHtml) { Remove-Item $tmpHtml -Force }
    }
}

function Get-HashArchivo {
    param([Parameter(Mandatory)][string]$Ruta)
    (Get-FileHash -LiteralPath $Ruta -Algorithm SHA256).Hash
}

$fallos = 0
foreach ($nombre in $Docs) {
    Write-Host "==> $nombre.pdf" -ForegroundColor Cyan
    Invoke-GenerarPdf -Nombre $nombre -Salida "$nombre.pdf"
    $hash = Get-HashArchivo -Ruta "$nombre.pdf"

    if ($Verify) {
        $tmp = "_$nombre.verify.pdf"
        Invoke-GenerarPdf -Nombre $nombre -Salida $tmp
        $hash2 = Get-HashArchivo -Ruta $tmp
        Remove-Item $tmp -Force

        if ($hash -eq $hash2) {
            Write-Host "    determinista OK  ($($hash.Substring(0,16)))" -ForegroundColor Green
        }
        else {
            Write-Host "    NO DETERMINISTA  $($hash.Substring(0,16)) != $($hash2.Substring(0,16))" -ForegroundColor Red
            $fallos++
        }
    }
    else {
        Write-Host "    sha256 $($hash.Substring(0,16))" -ForegroundColor DarkGray
    }
}

if ($fallos -gt 0) {
    Write-Host ""
    Write-Host "Atencion: $fallos PDF(s) no deterministas." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Listo. $($Docs.Count) PDF(s) generados." -ForegroundColor Green
