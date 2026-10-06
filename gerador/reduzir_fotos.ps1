# Reduz imagens para a web (micro-01, sem Python). Lê uma lista JSON escrita por gerar_mapa_produtos.py.
# Dois usos:
#   pwsh -File gerador\reduzir_fotos.ps1                       -> fotos dos cards: gerador\fotos_lista.json
#        (origem relativa a Fundo_infinito) -> fotos\<slug>.jpg (600 px) e fotos\g\<slug>.jpg (1600 px) no repositório
#   pwsh -File gerador\reduzir_fotos.ps1 -Celulas              -> miniaturas dos quadros: gerador\celulas_lista.json
#        (origem relativa a "público Catálogo", com destino e px por item) -> público Catálogo\mapa_produtos\celulas\
# Só refaz o que mudou (-Forcar refaz tudo).
param(
  [switch]$Celulas, [switch]$Forcar,
  [string]$Catalogo = (Join-Path $env:USERPROFILE "OneDrive - G MOVEIS LTDA\Dados - Documentos\gm\Imagens\público Catálogo"),
  [int]$Card = 600, [int]$Grande = 1600, [int]$Qualidade = 82
)
Add-Type -AssemblyName System.Drawing
$raiz = Split-Path $PSScriptRoot -Parent
$codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq "image/jpeg" }
$par = New-Object System.Drawing.Imaging.EncoderParameters(1)
$par.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]$Qualidade)

function Reduzir($img, $max, $destino) {
  $w = $img.Width; $h = $img.Height; $f = [Math]::Min(1.0, $max / [Math]::Max($w, $h))
  $nw = [int][Math]::Round($w * $f); $nh = [int][Math]::Round($h * $f)
  $bmp = New-Object System.Drawing.Bitmap($nw, $nh)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear([System.Drawing.Color]::White)   # PNG com transparência vira fundo branco
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
  $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.DrawImage($img, 0, 0, $nw, $nh)
  New-Item -ItemType Directory -Force -Path (Split-Path $destino -Parent) | Out-Null
  $bmp.Save($destino, $codec, $par)
  $g.Dispose(); $bmp.Dispose()
}

if ($Celulas) {
  $lista = Get-Content (Join-Path $PSScriptRoot "celulas_lista.json") -Raw -Encoding utf8 | ConvertFrom-Json
  $base = $Catalogo; $destRaiz = Join-Path $Catalogo "mapa_produtos\celulas"
} else {
  $lista = Get-Content (Join-Path $PSScriptRoot "fotos_lista.json") -Raw -Encoding utf8 | ConvertFrom-Json
  $base = Join-Path $Catalogo "Fundo_infinito"; $destRaiz = Join-Path $raiz "fotos"
}
$feitas = 0; $puladas = 0; $erros = @()
foreach ($e in $lista) {
  $orig = Join-Path $base $e.origem
  if (-not (Test-Path -LiteralPath $orig)) { $erros += "não achei: $($e.origem)"; continue }
  $mt = (Get-Item -LiteralPath $orig).LastWriteTimeUtc
  if ($Celulas) { $alvos = @(@{px = $e.px; dest = (Join-Path $destRaiz $e.destino)}) }
  else { $alvos = @(@{px = $Grande; dest = (Join-Path $destRaiz ("g\" + $e.slug + ".jpg"))}, @{px = $Card; dest = (Join-Path $destRaiz ($e.slug + ".jpg"))}) }
  $pronto = -not $Forcar
  foreach ($a in $alvos) { if (-not (Test-Path -LiteralPath $a.dest) -or (Get-Item -LiteralPath $a.dest).LastWriteTimeUtc -lt $mt) { $pronto = $false } }
  if ($pronto) { $puladas++; continue }
  try {
    $img = [System.Drawing.Image]::FromFile($orig)
    foreach ($a in $alvos) { Reduzir $img $a.px $a.dest }
    $img.Dispose(); $feitas++
  } catch { $erros += "erro em $($e.origem): $($_.Exception.Message)" }
}
"imagens reduzidas: $feitas · já estavam prontas: $puladas · erros: $($erros.Count)"
$erros | ForEach-Object { "  $_" }
