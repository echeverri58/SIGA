# Convertidor xlsx -> pdf en modo SERVIDOR (mantiene Excel abierto para ir rapido).
# Protocolo por stdin/stdout:
#   -> escribe:  <ruta.xlsx>|<ruta.pdf>
#   <- responde: OK   o   ERR <mensaje>
#   -> escribe:  QUIT   para terminar
$ErrorActionPreference = "Continue"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
$excel.Calculation = -4135          # xlCalculationManual (no recalcular: mas rapido)
$excel.Interactive = $false

[Console]::Out.WriteLine("READY")
[Console]::Out.Flush()

while ($true) {
  $line = [Console]::In.ReadLine()
  if ($null -eq $line) { break }
  $line = $line.Trim()
  if ($line -eq "") { continue }
  if ($line -eq "QUIT") { break }

  $parts = $line.Split("|")
  if ($parts.Length -lt 2) {
    [Console]::Out.WriteLine("ERR formato invalido")
    [Console]::Out.Flush()
    continue
  }

  $xlsx = $parts[0]
  $pdf  = $parts[1]
  $wb = $null
  try {
    # OJO: Workbooks.Open debe llamarse con UN solo argumento.
    # Con dos o mas argumentos la llamada COM falla en algunas versiones de Excel.
    $wb = $excel.Workbooks.Open($xlsx)
    $ws = $wb.Worksheets.Item("asistencia_mod")

    # --- Ajuste de impresion: la planilla completa en una hoja horizontal ---
    $ps = $ws.PageSetup
    $ps.PrintArea          = '$A$1:$M$33'
    $ps.Orientation        = 2      # xlLandscape
    $ps.Zoom               = $false
    $ps.FitToPagesWide     = 1
    $ps.FitToPagesTall     = 1
    $ps.CenterHorizontally = $true
    $ps.CenterVertically   = $false
    $ps.LeftMargin         = 12
    $ps.RightMargin        = 12
    $ps.TopMargin          = 12
    $ps.BottomMargin       = 12
    $ps.HeaderMargin       = 0
    $ps.FooterMargin       = 0

    $ws.ExportAsFixedFormat(0, $pdf)       # 0 = xlTypePDF (calidad estandar)
    $wb.Close($false)
    [Console]::Out.WriteLine("OK")
  }
  catch {
    try { if ($wb -ne $null) { $wb.Close($false) } } catch { }
    $msg = $_.Exception.Message -replace "[\r\n]+", " "
    [Console]::Out.WriteLine("ERR " + $msg)
  }
  [Console]::Out.Flush()
}

try { $excel.Quit() } catch { }
try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
