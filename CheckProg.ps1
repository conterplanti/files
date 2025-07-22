# Definir o caminho para salvar os arquivos
$OutputPath = "C:\Temp" # Você pode mudar para outro local, por exemplo: "$env:USERPROFILE\Documents\RelatoriosSoftware"

# Criar a pasta de saída se ela não existir
If (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory | Out-Null
}

# Obter o nome do computador e a data atual formatada
$ComputerName = $env:COMPUTERNAME
$CurrentDate = Get-Date -Format "yyyyMMdd_HHmm"

Write-Host "Iniciando a geração do relatório de softwares para '$ComputerName' em '$CurrentDate'..."
Write-Host "Os arquivos serão salvos em: $OutputPath"
Write-Host "----------------------------------------------------" # Linha divisória

Write-Host "`n-- Gerando relatório de Softwares Atualmente Instalados..."

# 1. Relatório de Softwares Atualmente Instalados (como em "Aplicativos e Recursos")
$InstalledSoftwarePath = Join-Path $OutputPath "SoftwaresInstalados_Geral_${ComputerName}_${CurrentDate}.csv"
Try {
    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
      Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
      Where-Object {$_.DisplayName -ne $null} |
      Sort-Object DisplayName |
      Export-Csv -Path $InstalledSoftwarePath -NoTypeInformation -Encoding UTF8 -Force

    Write-Host "Relatório de softwares instalados salvo em: $InstalledSoftwarePath"
}
Catch {
    Write-Warning "Erro ao gerar o relatório de softwares instalados: $($_.Exception.Message)"
}

Write-Host "----------------------------------------------------" # Linha divisória
Write-Host "`n-- Gerando relatório de Histórico de Instalação/Desinstalação (Logs de Eventos)..."

# 2. Relatório de Histórico de Instalação/Desinstalação (via Logs de Eventos)
$EventHistoryPath = Join-Path $OutputPath "HistoricoSoftware_EventLogs_${ComputerName}_${CurrentDate}.csv"
Try {
    $Events = Get-WinEvent -LogName Application -FilterXPath '*[System[(Provider[@Name="MsiInstaller"]) and (EventID=1033 or EventID=1034 or EventID=11707 or EventID=11724)]]' -MaxEvents 10000 | # Aumentado para 10000 eventos para um histórico maior
      Select-Object TimeCreated, Id, Message |
      ForEach-Object {
        [PSCustomObject]@{
          Timestamp = $_.TimeCreated
          EventID   = $_.Id
          EventType = switch ($_.Id) {
            1033  {"Installation (Product Name)"}
            1034  {"Uninstallation (Product Name)"}
            11707 {"Installation (Successful)"}
            11724 {"Uninstallation (Successful)"}
            default {"Unknown"}
          }
          Message   = $_.Message
        }
      }

    If ($Events) {
        $Events | Export-Csv -Path $EventHistoryPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "Relatório de histórico de eventos salvo em: $EventHistoryPath"
    } Else {
        Write-Host "Nenhum evento de instalação/desinstalação (MsiInstaller) encontrado nos logs recentes."
    }
}
Catch {
    Write-Warning "Erro ao gerar o relatório de histórico de eventos: $($_.Exception.Message)"
}

Write-Host "----------------------------------------------------" # Linha divisória
Write-Host "`nProcesso concluído."
