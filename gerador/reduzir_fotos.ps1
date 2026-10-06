# Reduz as fotos do catálogo para a web (micro-01, sem Python): lê gerador\fotos_lista.json
# (origem -> slug, escrito por gerar_mapa_produtos.py) e grava fotos\<slug>.jpg (600 px)
# e fotos\g\<slug>.jpg (1600 px) na raiz do repositório. Só refaz o que mudou.
# Uso: pwsh -File gerador\reduzir_fotos.ps1   (de qualquer pasta)
param(
  [string]$Fundo = (Join-Path $env:USERPROFILE "OneDrive - G MOVEIS LTDA\Dados - Documentos\gm\Imagens\público Catálogo\Fundo_infinito"),
  [int]$Card = 600, [int]$Grande = 1600, [int]$Qualidade = 82, [switch]$Forcar
)
Add-Type -AssemblyName System.Drawing
$raiz = Split-Path $PSScriptRoot -Parent
$lista = Get-Content (Join-Path $PSScriptRoot "fotos_lista.json") -Raw -Encoding utf8 | ConvertFrom-Json
New-Item -ItemType Directory -Force -Path (Join-Path $raiz "fotos"), (Join-Path $raiz "fotos\g") | Out-Null
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
  $bmp.Save($destino, $codec, $par)
  $g.Dispose(); $bmp.Dispose()
}

$feitas = 0; $puladas = 0; $erros = @()
foreach ($e in $lista) {
  $orig = Join-Path $Fundo $e.origem
  $d1 = Join-Path $raiz ("fotos\" + $e.slug + ".jpg")
  $d2 = Join-Path $raiz ("fotos\g\" + $e.slug + ".jpg")
  if (-not (Test-Path -LiteralPath $orig)) { $erros += "não achei: $($e.origem)"; continue }
  $mt = (Get-Item -LiteralPath $orig).LastWriteTimeUtc
  if (-not $Forcar -and (Test-Path -LiteralPath $d1) -and (Test-Path -LiteralPath $d2) -and (Get-Item -LiteralPath $d1).LastWriteTimeUtc -ge $mt) { $puladas++; continue }
  try {
    $img = [System.Drawing.Image]::FromFile($orig)
    Reduzir $img $Grande $d2
    Reduzir $img $Card $d1
    $img.Dispose(); $feitas++
  } catch { $erros += "erro em $($e.origem): $($_.Exception.Message)" }
}
"fotos reduzidas: $feitas · já estavam prontas: $puladas · erros: $($erros.Count)"
$erros | ForEach-Object { "  $_" }
