# Genera Siga_planilla.xlsx: copia de Siga.xlsx con SOLO la hoja "asistencia_mod".
# Se usa para que la conversion a PDF (Excel o LibreOffice) produzca solo la planilla.
#
# IMPORTANTE: la plantilla se genera con Excel. Armarla por fuera de Excel produce
# un archivo que Excel rechaza al abrirlo ("No se puede obtener la propiedad Open").
#
# Uso:  powershell -NoProfile -ExecutionPolicy Bypass -File crear_plantilla_pdf.ps1
$ErrorActionPreference = "Stop"

$dir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$origen = Join-Path $dir "Siga.xlsx"
$destino = Join-Path $dir "Siga_planilla.xlsx"
$otras = @("asistencia_charlas_informativas", "Instrucciones", "lista elegibles")

if (-not (Test-Path $origen)) { Write-Error "No se encontro Siga.xlsx en $dir"; exit 1 }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($origen)
  foreach ($n in $otras) {
    try { $wb.Worksheets.Item($n).Delete() } catch { }
  }
  if (Test-Path $destino) { Remove-Item $destino -Force }
  $wb.SaveAs($destino, 51)      # 51 = xlOpenXMLWorkbook (.xlsx)
  $wb.Close($false)
  Write-Output "Creado: $destino"
}
finally {
  try { $excel.Quit() } catch { }
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
