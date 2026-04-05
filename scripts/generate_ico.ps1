Add-Type -AssemblyName System.Drawing
$imgPath = "E:\Code\QSBar\Release\Installinfo\logo_icon.png"
$icoPath = "E:\Code\QSBar\Installinfo\qsbar.ico"

$srcImg = [System.Drawing.Image]::FromFile($imgPath)
$size = 64
$bmp = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.DrawImage($srcImg, 0, 0, $size, $size)
$g.Dispose()

$icoFile = [System.IO.File]::Create($icoPath)

$icoFile.WriteByte(0); $icoFile.WriteByte(0)
$icoFile.WriteByte(1); $icoFile.WriteByte(0)
$icoFile.WriteByte(1); $icoFile.WriteByte(0)

$icoFile.WriteByte($size); $icoFile.WriteByte($size)
$icoFile.WriteByte(0); $icoFile.WriteByte(0)
$icoFile.WriteByte(1); $icoFile.WriteByte(0)
$icoFile.WriteByte(32); $icoFile.WriteByte(0)

$andMaskStride = [math]::Floor(($size + 31) / 32) * 4
$andMaskSize = $andMaskStride * $size
$imageDataSize = 40 + ($size * $size * 4) + $andMaskSize

$icoFile.Write([System.BitConverter]::GetBytes([int]$imageDataSize), 0, 4)
$icoFile.Write([System.BitConverter]::GetBytes([int]22), 0, 4)

$icoFile.Write([System.BitConverter]::GetBytes([int]40), 0, 4)
$icoFile.Write([System.BitConverter]::GetBytes([int]$size), 0, 4)
$icoFile.Write([System.BitConverter]::GetBytes([int]($size * 2)), 0, 4)
$icoFile.Write([System.BitConverter]::GetBytes([short]1), 0, 2)
$icoFile.Write([System.BitConverter]::GetBytes([short]32), 0, 2)
$icoFile.Write([System.BitConverter]::GetBytes([int]0), 0, 4)
$icoFile.Write([System.BitConverter]::GetBytes([int]($size * $size * 4)), 0, 4)
$icoFile.Write([System.BitConverter]::GetBytes([int]0), 0, 4)
$icoFile.Write([System.BitConverter]::GetBytes([int]0), 0, 4)
$icoFile.Write([System.BitConverter]::GetBytes([int]0), 0, 4)
$icoFile.Write([System.BitConverter]::GetBytes([int]0), 0, 4)

for ($y = $size - 1; $y -ge 0; $y--) {
    for ($x = 0; $x -lt $size; $x++) {
        $c = $bmp.GetPixel($x, $y)
        $icoFile.WriteByte($c.B)
        $icoFile.WriteByte($c.G)
        $icoFile.WriteByte($c.R)
        $icoFile.WriteByte($c.A)
    }
}

$andMask = New-Object byte[] $andMaskSize
$icoFile.Write($andMask, 0, $andMask.Length)

$icoFile.Close()
$bmp.Dispose()
$srcImg.Dispose()
Write-Host "True 64x64 uncompressed Icon created successfully!"