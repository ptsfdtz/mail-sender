param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot "..\src-tauri\icons")
)

Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null

$size = 256
$bitmap = [System.Drawing.Bitmap]::new($size, $size)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.Clear([System.Drawing.Color]::FromArgb(23, 32, 29))

$amber = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(234, 183, 103))
$darkPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(23, 32, 29), 14)
$darkPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
$graphics.FillRectangle($amber, 43, 67, 170, 122)
$graphics.DrawRectangle($darkPen, 43, 67, 170, 122)
$graphics.DrawLine($darkPen, 48, 74, 128, 139)
$graphics.DrawLine($darkPen, 208, 74, 128, 139)
$graphics.DrawLine($darkPen, 48, 183, 104, 128)
$graphics.DrawLine($darkPen, 208, 183, 152, 128)

$pngPath = Join-Path $OutputDirectory "icon.png"
$bitmap.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$darkPen.Dispose()
$amber.Dispose()
$bitmap.Dispose()

$pngBytes = [System.IO.File]::ReadAllBytes($pngPath)
$stream = [System.IO.MemoryStream]::new()
$writer = [System.IO.BinaryWriter]::new($stream)
$writer.Write([uint16]0)
$writer.Write([uint16]1)
$writer.Write([uint16]1)
$writer.Write([byte]0)
$writer.Write([byte]0)
$writer.Write([byte]0)
$writer.Write([byte]0)
$writer.Write([uint16]1)
$writer.Write([uint16]32)
$writer.Write([uint32]$pngBytes.Length)
$writer.Write([uint32]22)
$writer.Write($pngBytes)
[System.IO.File]::WriteAllBytes((Join-Path $OutputDirectory "icon.ico"), $stream.ToArray())
$writer.Dispose()
$stream.Dispose()
