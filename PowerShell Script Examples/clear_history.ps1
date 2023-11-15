$havePSReadline = ($null -ne (Get-Module -EA SilentlyContinue PSReadline))

if ($havePSReadline) {
  Clear-Host
  if (Test-Path (Get-PSReadlineOption).HistorySavePath) { 
    Remove-Item -EA Stop (Get-PSReadlineOption).HistorySavePath 
    $null = New-Item -Type File -Path (Get-PSReadlineOption).HistorySavePath
  }
  Clear-History
  [Microsoft.PowerShell.PSConsoleReadLine]::ClearHistory()
} else {
  Clear-Host
  $null = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
  [System.Windows.Forms.SendKeys]::SendWait('%{F7 2}')
  Clear-History
}
