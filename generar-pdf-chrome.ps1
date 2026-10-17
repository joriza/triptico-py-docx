<#
.SYNOPSIS
    Genera los PDFs del kit "Tríptico para un emprendimiento casero de gelatina"
    a partir de los .md, con Chrome headless (motor Blink), de forma DETERMINÍSTICA.

.DESCRIPTION
    Pipeline:
      1. pandoc : <doc>.md -> _gen.tmp.html (con custom-markdown-chrome.css)
      2. chrome : _gen.tmp.html -> <doc>.pdf
      3. normalizar: reemplaza las fechas del PDF por una constante

    POR QUÉ CHROME Y NO WKHTMLTOPDF
    wkhtmltopdf 0.12.6 usa QtWebKit 4.8 (2011) y desobedece la plantilla en cuatro
    cosas, todas medidas sobre el PDF resultante:
      - escala las longitudes del CSS ~0,80 (declarado != impreso);
      - ignora @page: ni el margen ni el pie declarados llegan al papel;
      - ignora line-height: declara 1,5 y dibuja ~1,1;
      - ignora column-count y no hace fallback por glifo (de ahí el truco de
        poner "Segoe UI Emoji" al frente de font-family).
    Blink (Chrome/Edge) respeta las cuatro. Por eso la plantilla de este pipeline
    es custom-markdown-chrome.css y no custom-markdown-v2.css: son hermanas, pero
    cada una está escrita para su motor, y mezclarlas da un PDF mal maquetado.

    POR QUÉ ES NECESARIO EL PASO 3
    Chrome escribe CreationDate y ModDate en los metadatos: DOS fechas, no una
    (medido: offsets 248 y 283 del archivo). Sin normalizar, dos ejecuciones con
    la misma entrada dan archivos distintos. El paso 3 reemplaza las dos por una
    constante de la MISMA longitud (23 caracteres), lo que preserva todos los
    offsets internos del PDF. No se detectó ninguna otra fecha variable (no hay
    XMP ni /ID aleatorio).
    No cambiar la fecha por una de distinta longitud: desplazaría los offsets de
    la tabla xref y corrompería el archivo.

    CANTIDAD DE COLUMNAS
    Regla desde 2026-10-17: UNA columna por defecto en todos los documentos;
    dos columnas solo a pedido explícito con -Columnas 2. La razón de fondo no
    cambió: el texto corrido gana legibilidad a dos columnas (el largo de línea
    baja de ~105 a ~41-46 caracteres), pero una tabla partida o apretada en una
    columna angosta se lee mal. El $ColumnasPorDoc conserva la proporción de
    tabla medida sobre cada .md por si algún documento vuelve a dos columnas.
    Los elementos de ancho completo (título principal y tablas) llevan
    column-span: all en el CSS, así que no se parten entre columnas.

    NOTAS DE MAQUETADO (medidas sobre el PDF resultante)
    - Los márgenes de página salen del @page del CSS, 1:1 (2,5 cm declarados dan
      2,49 cm impresos). Se pueden pisar con -MargenLateral, que reescribe el
      valor en el CSS temporal: el número que se pasa es el número que se imprime.
    - El pie lo arma el CSS con contadores de página. No se usan los flags
      --footer-* de wkhtmltopdf.
    - El encabezado con el nombre del documento (sin extensión) lo escribe el
      script en el CSS temporal, en un bloque @top-center: el CSS no puede leer el
      nombre del archivo, porque Blink no implementa string-set/string() (medido:
      el box queda vacío). No agrega papel: el encabezado y el pie viven DENTRO
      del margen y no le roban ancho al texto (medido con la configuración
      previa: el kit da 29 hojas con y sin encabezado). Se apaga con -SinEncabezado.
    - La fecha de generación va en el pie, en la esquina del lomo (abajo a la
      izquierda en la hoja impar, abajo a la derecha en la par; en modo
      simétrico, abajo a la izquierda), en ISO-8601 y solo la fecha. La inyecta
      el script: Blink no implementa string-set. No rompe el determinismo: dos
      corridas del mismo día dan bytes idénticos (la fecha es parte de la
      entrada, como el nombre del archivo).
    - Los márgenes vienen ESPEJADOS por defecto (ver -MargenLomo): el kit termina
      junto en un anillado, así que el lomo tiene que caer siempre del mismo lado
      en los cuatro documentos, aunque alguno sea de una sola hoja. Una sola forma
      de regenerar da un juego encuadernable en todos los casos.
    - El script se puede invocar desde cualquier carpeta: los temporales y la
      salida se resuelven contra la carpeta del script. Con nombres relativos,
      .NET los resolvía contra el CWD del proceso, que no sigue a Set-Location, y
      el PDF podía salir sin estilos y sin aviso (medido: hoja Letter y márgenes
      del navegador en vez de A4 y 2,5 cm).
    - Blink trae márgenes propios de párrafo y de lista (1em arriba y abajo) que el
      WebKit viejo no aplicaba; el CSS los resetea a propósito. Si se quitan, cada
      documento crece alrededor de una hoja.
    - El HTML y el CSS temporales DEBEN quedar en esta misma carpeta: el CSS se
      enlaza con ruta relativa y Chrome resuelve href contra la ruta del HTML.

.PARAMETER Doc
    Nombres de documento sin la extensión .md. Si se omite, se generan todos.

.PARAMETER MargenLateral
    Activa el modo SIMÉTRICO: el mismo margen a los dos costados, para imprimir
    de un solo lado. Con este flag no se generan los bloques de márgenes
    espejados. Si se omite (el caso normal), los márgenes son espejados:
    -MargenLomo del lado del lomo y -MargenExterno del otro.

    Referencia de legibilidad: hasta 90 caracteres por línea es cómodo y ~66 es
    lo ideal. A una columna este margen es el único control del largo de línea;
    con -Columnas 2 el largo de línea queda en ~41 caracteres sin tocar el margen.

.PARAMETER MargenVertical
    Margen superior e inferior del PDF, con unidad (1.5cm, 15mm, 1in). Por
    defecto: 1.5cm, que es el valor declarado en el @page de la plantilla.

    Ojo: este valor NO es una herencia de un diseño anterior. El pipeline viejo
    (wkhtmltopdf) ignoraba el @page y usaba 10mm arriba y abajo por flag, así que
    los PDF del kit que estaban aprobados tienen 1 cm de margen vertical, no 2,5.
    Es el parámetro que más papel mueve, porque define el alto útil de cada hoja:
    8% del alto útil por cada cm que se agrega arriba y abajo.

    Medido sobre el kit completo (11 pt, guía y modelo a 2 columnas, cuaderno y
    rúbrica a 1), en hojas por documento:
        margen vertical  interlineado  guía  cuaderno  modelo  rúbrica  total
        1,5cm            1,3           10     10        4       5        29
        2,5cm            1,3           10     11        4       5        30
        1,5cm            1,5           11     11        4       5        31
        2,5cm            1,5           12     12        5       6        35
    Las dos primeras filas salen de un flag; las dos últimas piden además mover
    el interlineado a 1,5 (una línea del body de la plantilla). El largo de línea
    no cambia entre las cuatro: 41 caracteres en la guía y el modelo, 88 en el
    cuaderno, 56 en la rúbrica.

    Esas cuatro filas son del modo SIMÉTRICO (-MargenLateral), medidas con la
    configuración previa (guía y modelo a 2 columnas). El modo por defecto
    (espejado) con el estándar nuevo, lomo 1,5 cm y externo 1 cm, ensancha el
    texto a 18,5 cm; los conteos de arriba quedan como referencia del estándar
    anterior.

.PARAMETER MargenLomo
    Margen del lado del lomo (en la hoja impar, la izquierda; en la par, la
    derecha). Por defecto: 1.5cm, estándar desde 2026-10-17 (en documentos
    largos se ahorra papel). Los márgenes vienen ESPEJADOS por defecto:
    el kit se anilla junto, así que el lomo cae siempre del mismo lado en los
    cuatro documentos, aunque alguno sea de una sola hoja. Es la única forma de
    regenerar que da un juego encuadernable en todos los casos.

    Medido en Blink con lomo 2.5cm y externo 1cm (los defaults hasta el
    2026-10-17): hojas impares
    2,49 cm del lado del lomo y hojas pares 1,01 cm del lado externo. Ojo con
    esto: el ancho de texto es 21 cm menos lomo menos externo (17,5 cm en vez de
    los 16 del modo simétrico), así que el largo de línea y la cantidad de hojas
    cambian: el kit da 28 hojas contra 29 del simétrico (la guía pasa de 10 a 9,
    el resto queda igual).

    Para el modo simétrico (impresión de un solo lado) use -MargenLateral, que
    es excluyente con este parámetro.

.PARAMETER MargenExterno
    Margen del lado externo (el opuesto al lomo) en el modo espejado, que es el
    que viene por defecto. Por defecto: 1cm. No tiene sentido en modo simétrico.

.PARAMETER SinEncabezado
    No imprime el encabezado con el nombre del documento en el margen superior.
    Por defecto el encabezado va.

.PARAMETER Columnas
    Fuerza la cantidad de columnas para todos los documentos indicados. Si se
    omite, cada documento usa la suya de $ColumnasPorDoc.

.PARAMETER Verify
    Genera cada PDF dos veces y compara los hashes SHA256 para comprobar el
    determinismo. Tarda el doble.

.EXAMPLE
    .\generar-pdf-chrome.ps1                                  # los 4, margenes espejados
    .\generar-pdf-chrome.ps1 -Doc rubrica-triptico-gelatina
    .\generar-pdf-chrome.ps1 -Columnas 2                     # dos columnas, solo a pedido explicito
    .\generar-pdf-chrome.ps1 -Verify
    .\generar-pdf-chrome.ps1 -SinEncabezado
    .\generar-pdf-chrome.ps1 -MargenLateral 2.5cm             # modo simetrico (un solo lado)
    .\generar-pdf-chrome.ps1 -MargenLomo 2cm -MargenExterno 1.5cm
#>

[CmdletBinding()]
param(
    [string[]]$Doc,
    [string]$MargenLateral = '2.5cm',
    [string]$MargenVertical = '1.5cm',
    [string]$MargenLomo = '1.5cm',
    [string]$MargenExterno = '1cm',
    [int]$Columnas = 0,
    [switch]$SinEncabezado,
    [switch]$Verify
)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

# El modo por defecto son márgenes ESPEJADOS. El kit termina junto en un anillado,
# así que el lomo tiene que caer siempre del mismo lado en los cuatro documentos,
# aunque alguno sea de una sola hoja: es la única forma de que una regeneración
# única dé un juego encuadernable. -MargenLateral es la salida explícita al modo
# simétrico, para la impresión de un solo lado.
$simetrico = $PSBoundParameters.ContainsKey('MargenLateral') -and -not $PSBoundParameters.ContainsKey('MargenLomo')

$validos = @(@('MargenLateral', $MargenLateral), @('MargenVertical', $MargenVertical), @('MargenLomo', $MargenLomo), @('MargenExterno', $MargenExterno))
foreach ($par in $validos) {
    if ($par[1] -notmatch '^\d+(\.\d+)?(cm|mm|in|pt|px)$') {
        throw "$($par[0]) '$($par[1])' no tiene un formato válido. Ejemplos: 2.5cm, 25mm, 1in."
    }
}
if ($PSBoundParameters.ContainsKey('MargenLateral') -and $PSBoundParameters.ContainsKey('MargenLomo')) {
    throw "-MargenLateral y -MargenLomo son excluyentes: con márgenes espejados el ancho del texto sale de -MargenLomo y -MargenExterno, y -MargenLateral quedaría ignorado en silencio."
}
if ($simetrico -and $PSBoundParameters.ContainsKey('MargenExterno')) {
    throw "-MargenExterno solo tiene sentido con márgenes espejados, que es el modo por defecto: para el modo simétrico use solo -MargenLateral."
}
if ($Columnas -lt 0 -or $Columnas -gt 3) {
    throw "Columnas tiene que estar entre 1 y 3 (o 0 para usar el valor por documento)."
}

$CSS      = 'custom-markdown-chrome.css'
$TodosDoc = @(
    'guia-docente-triptico-gelatina',
    'cuaderno-del-alumno-triptico-gelatina',
    'modelo-triptico-gelatina',
    'rubrica-triptico-gelatina'
)
$Docs = if ($PSBoundParameters.ContainsKey('Doc')) { $Doc } else { $TodosDoc }

# Columnas por documento. Regla desde 2026-10-17: UNA columna en todos los
# documentos; dos columnas solo a pedido explícito (-Columnas 2). El mapa
# conserva la proporción de tabla medida sobre cada .md (líneas que empiezan
# con "|"; las listas cortas y densas cuentan como tabla a los fines de la
# lectura, porque un ítem de dos palabras por línea se lee peor en una columna
# angosta que un párrafo) por si algún documento vuelve a maquetarse a dos.
$ColumnasPorDoc = @{
    'guia-docente-triptico-gelatina'         = 1   # 11% tabla, prosa
    'cuaderno-del-alumno-triptico-gelatina'  = 1   # 16% tabla + listas densas
    'modelo-triptico-gelatina'               = 1   #  5% tabla, prosa
    'rubrica-triptico-gelatina'              = 1   # 32% tabla, la más tabular
}

# Ruta del ejecutable de Chrome. Se prueban las dos ubicaciones normales y Edge
# como último recurso (Blink en los tres casos).
$Chrome = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $Chrome) { throw "No se encontró Chrome ni Edge." }

# Perfil de Chrome fuera de la carpeta del proyecto, para no ensuciarla.
$Perfil = Join-Path $env:TEMP 'gentle-ai-chrome-perfil'
New-Item -ItemType Directory -Path $Perfil -Force | Out-Null

# Fecha fija, 23 caracteres: D:YYYYMMDDHHmmSS + HH'MM'
$FechaFija   = "D:20260101000000+00'00'"
$PatronFecha = "D:\d{14}[+-]\d{2}'\d{2}'"

if (-not (Get-Command 'pandoc' -ErrorAction SilentlyContinue)) {
    throw "Falta 'pandoc' en el PATH."
}

function Set-FechaDeterministica {
    param([Parameter(Mandatory)][string]$Ruta)

    # Latin1 (ISO-8859-1) mapea 1:1 cada byte, asi que sirve para tratar el PDF
    # como texto sin romper los bytes binarios.
    $enc   = [System.Text.Encoding]::GetEncoding(28591)
    $bytes = [System.IO.File]::ReadAllBytes($Ruta)
    $texto = $enc.GetString($bytes)
    $todas = [regex]::Matches($texto, $PatronFecha)

    # Chrome escribe dos: CreationDate y ModDate.
    if ($todas.Count -lt 1) {
        throw "$Ruta`: no se encontro ninguna fecha para normalizar. El PDF no seria determinista."
    }

    $nuevo = $texto
    foreach ($m in ($todas | Sort-Object Index -Descending)) {
        if ($m.Value.Length -ne $FechaFija.Length) {
            throw "$Ruta`: la fecha '$($m.Value)' mide $($m.Value.Length) y la fija mide $($FechaFija.Length). Igualar longitudes antes de continuar: un reemplazo de otra longitud desplaza los offsets del PDF."
        }
        $nuevo = $nuevo.Substring(0, $m.Index) + $FechaFija + $nuevo.Substring($m.Index + $m.Value.Length)
    }
    [System.IO.File]::WriteAllBytes($Ruta, $enc.GetBytes($nuevo))
}

function New-CssTemporal {
    param(
        [Parameter(Mandatory)][int]$Columnas,
        [Parameter(Mandatory)][string]$Nombre,
        [Parameter(Mandatory)][string]$Destino
    )

    $css = Get-Content -LiteralPath $CSS -Raw

    # Los dos parches van por regex y verifican que hayan entrado. Un reemplazo
    # que no matchea deja el PDF con los valores de la plantilla y el flag
    # parecería ignorado en silencio: exactamente el defecto que se está dejando
    # atrás, así que si no matchea se corta.
    $reColumnas = 'column-count:\s*\d+;'
    $n = ([regex]::Matches($css, $reColumnas)).Count
    if ($n -ne 1) {
        throw "En $CSS se esperaba 1 declaración 'column-count', se encontraron $n. Revisar la plantilla antes de parchearla."
    }
    $css = [regex]::Replace($css, $reColumnas, "column-count: $Columnas;")

    # El margen se busca DENTRO de @page: hay otros 'margin' en la hoja de
    # estilos (los de cada título) que no se tienen que tocar.
    $reMargen = '(?s)(@page\s*\{.*?)margin:\s*[^;]+;'
    $n = ([regex]::Matches($css, $reMargen)).Count
    if ($n -ne 1) {
        throw "En $CSS se esperaba 1 margen dentro de @page, se encontraron $n. Revisar la plantilla antes de parchearla."
    }
    # El @page base toma la configuración de la hoja impar (lomo a la izquierda):
    # es la que vale si algún motor no resolviera los bloques :left/:right. En
    # modo simétrico queda el margen de los dos lados igual.
    $margen = if ($simetrico) {
        "$MargenVertical $MargenLateral"
    }
    else {
        "$MargenVertical $MargenExterno $MargenVertical $MargenLomo"
    }
    $css = [regex]::Replace($css, $reMargen, "`${1}margin: $margen;")

    # Encabezado con el nombre del documento. Lo inyecta el script y no la
    # plantilla porque el CSS no puede leer el nombre del archivo: Blink no
    # implementa string-set/string() (medido, el box queda vacío). El nombre es
    # el mismo que va en --variable=pagetitle de pandoc.
    $reEncabezado = '\/\*\[\[ENCABEZADO\]\]\*\/'
    $n = ([regex]::Matches($css, $reEncabezado)).Count
    if ($n -ne 1) {
        throw "En $CSS se esperaba 1 marca de encabezado, se encontraron $n. Revisar la plantilla antes de parchearla."
    }
    $encabezado = ''
    if (-not $SinEncabezado) {
        $texto = $Nombre -replace '"', '\"'
        $encabezado = "@top-center {`n    content: `"$texto`";`n    font-size: 8pt;`n    color: #555555;`n    font-family: Arial, Helvetica, sans-serif;`n  }"
    }
    $css = [regex]::Replace($css, $reEncabezado, $encabezado)

    # Márgenes espejados: se escriben salvo que se pida el modo simétrico.
    $reEspejado = '\/\*\[\[ESPEJADO\]\]\*\/'
    $n = ([regex]::Matches($css, $reEspejado)).Count
    if ($n -ne 1) {
        throw "En $CSS se esperaba 1 marca de márgenes espejados, se encontraron $n. Revisar la plantilla antes de parchearla."
    }

    # Fecha ISO-8601 (solo la fecha) en la esquina del lomo del pie. La inyecta
    # el script y no la plantilla porque el CSS no puede leer el reloj, igual
    # que pasa con el nombre del documento en el encabezado. Dos corridas el
    # MISMO dia producen los mismos bytes: el determinismo se conserva, la fecha
    # es parte de la entrada, igual que el nombre del archivo.
    $fechaPie = Get-Date -Format 'yyyy-MM-dd'
    $estiloFecha = "`n    font-size: 8pt;`n    color: #555555;`n    font-family: Arial, Helvetica, sans-serif;`n  }"
    $bloqueFechaImpar = "@bottom-left {`n    content: `"$fechaPie`";$estiloFecha"
    $bloqueFechaPar   = "@bottom-right {`n    content: `"$fechaPie`";$estiloFecha"

    $espejado = ''
    if (-not $simetrico) {
        $espejado = @"
/* Hoja impar (la 1 lo es): lomo a la izquierda. Hoja par: lomo a la derecha.
   La fecha va en la esquina del lomo: abajo a la izquierda en la impar y
   abajo a la derecha en la par. */
@page :right {
  margin: $MargenVertical $MargenExterno $MargenVertical $MargenLomo;
  $bloqueFechaImpar
}
@page :left  {
  margin: $MargenVertical $MargenLomo $MargenVertical $MargenExterno;
  $bloqueFechaPar
}
"@
    }
    else {
        # Modo simetrico (impresion de un solo lado): no hay lomo, la fecha va
        # abajo a la izquierda como en una hoja impar. Es un @page adicional y
        # Blink lo fusiona con el @page base de la plantilla.
        $espejado = @"
/* Modo simetrico: sin espejado. La fecha queda abajo a la izquierda. */
@page {
  $bloqueFechaImpar
}
"@
    }
    $css = [regex]::Replace($css, $reEspejado, $espejado)

    # Con una sola columna, llenar en orden evita el hueco que deja el balanceo.
    if ($Columnas -eq 1) {
        $css = [regex]::Replace($css, 'column-gap:\s*[^;]+;', "column-gap: 0.9cm;`n  column-fill: auto;")
    }

    # Sin BOM: el CSS se enlaza tal cual y un BOM al principio de un archivo CSS
    # no aporta nada.
    [System.IO.File]::WriteAllText($Destino, $css, (New-Object System.Text.UTF8Encoding($false)))
}

function Invoke-GenerarPdf {
    param(
        [Parameter(Mandatory)][string]$Nombre,
        [Parameter(Mandatory)][string]$Salida
    )

    if (-not (Test-Path (Join-Path $PSScriptRoot "$Nombre.md"))) { throw "No existe $Nombre.md" }

    $col = if ($Columnas -gt 0) { $Columnas } else { $ColumnasPorDoc[$Nombre] }
    if (-not $col) {
        Write-Warning "$Nombre no está en el mapa de columnas: se usa 1."
        $col = 1
    }

    # El HTML y el CSS temporales van en ESTA carpeta para que el CSS relativo
    # resuelva (ver la nota de maquetado del encabezado).
    # Con ruta ABSOLUTA, no relativa: [System.IO.File]::WriteAllText resuelve los
    # nombres relativos contra el CWD del proceso, que NO sigue a Set-Location.
    # Invocando el script desde otra carpeta, el CSS quedaba en esa otra carpeta,
    # Chrome no lo encontraba y el PDF salía SIN estilos y sin aviso (medido:
    # hoja Letter y márgenes del navegador en vez de A4 y 2,5 cm). El nombre
    # relativo se usa solo para el --css de pandoc, que va tal cual al href del
    # HTML y por eso tiene que ser relativo.
    $tmpCssNombre  = '_gen.tmp.css'
    $tmpHtmlNombre = '_gen.tmp.html'
    $tmpCss  = Join-Path $PSScriptRoot $tmpCssNombre
    $tmpHtml = Join-Path $PSScriptRoot $tmpHtmlNombre
    try {
        New-CssTemporal -Columnas $col -Nombre $Nombre -Destino $tmpCss

        pandoc (Join-Path $PSScriptRoot "$Nombre.md") `
            --from=markdown+task_lists+hard_line_breaks `
            --standalone `
            --css=$tmpCssNombre `
            --variable="pagetitle:$Nombre" `
            -o $tmpHtml

        $url  = 'file:///' + ((Resolve-Path -LiteralPath $tmpHtml).Path -replace '\\', '/')
        $dest = Join-Path $PSScriptRoot $Salida

        & $Chrome --headless=new --disable-gpu --no-pdf-header-footer `
            --user-data-dir="$Perfil" --print-to-pdf="$dest" $url | Out-Null

        if (-not (Test-Path -LiteralPath $dest)) {
            throw "Chrome no generó $Salida (revisar la salida de Chrome)."
        }

        Set-FechaDeterministica -Ruta $dest
    }
    finally {
        if (Test-Path $tmpHtml) { Remove-Item $tmpHtml -Force }
        if (Test-Path $tmpCss)  { Remove-Item $tmpCss  -Force }
    }
}

function Get-HashArchivo {
    param([Parameter(Mandatory)][string]$Ruta)
    (Get-FileHash -LiteralPath $Ruta -Algorithm SHA256).Hash
}

$fallos = 0
foreach ($nombre in $Docs) {
    $col = if ($Columnas -gt 0) { $Columnas } else { $ColumnasPorDoc[$nombre] }
    Write-Host "==> $nombre.pdf  ($col columna(s))" -ForegroundColor Cyan
    Invoke-GenerarPdf -Nombre $nombre -Salida "$nombre.pdf"
    $hash = Get-HashArchivo -Ruta "$nombre.pdf"

    if ($Verify) {
        $tmp   = "_$nombre.verify.pdf"
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
