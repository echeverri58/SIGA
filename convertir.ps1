# Convierte un .xlsx a .pdf usando Microsoft Excel (COM).
# Uso: powershell -NoProfile -ExecutionPolicy Bypass -File convertir.ps1 <entrada.xlsx> <salida.pdf>
param(
  [Parameter(Mandatory = $true)][string]$ExcelPath,
  [Parameter(Mandatory = $true)][string]$PdfPath
)

$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

try {
  $wb = $excel.Workbooks.Open($ExcelPath)
  # Exportar SOLO la hoja de la planilla de asistencia
  $ws = $wb.Worksheets.Item("asistencia_mod")
  # Ajuste de impresion (si algo falla aqui, la conversion continua igual)
  try {
    $ps = $ws.PageSetup
    $ps.PrintArea          = '$A$1:$M$33'
    $ps.Orientation        = 2      # xlLandscape
    $ps.Zoom               = $false
    $ps.FitToPagesWide     = 1
    $ps.FitToPagesTall     = 1
    $ps.CenterHorizontally = $true
  } catch { }
  # xlTypePDF = 0
  $ws.ExportAsFixedFormat(0, $PdfPath)
  $wb.Close($false)
}
finally {
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
