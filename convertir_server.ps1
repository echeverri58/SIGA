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
    $wb = $excel.Workbooks.Open($xlsx, 0)  # 0 = no actualizar vinculos
    $ws = $wb.Worksheets.Item("asistencia_mod")
    $ws.ExportAsFixedFormat(0, $pdf)       # 0 = xlTypePDF
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
