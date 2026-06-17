# ============================================================
# Agendador de Tarefas - MG Contecnica Chamados (Extrator + Base + Push)
# ============================================================

$python = (Get-Command python).Source
$script = "C:\Users\gamaral\Desktop\Python\Chamados\backend\extrato_chamados.py"
$horarios = @("09:00", "13:00", "17:00", "18:00")

Write-Host "Configurando tarefas agendadas..." -ForegroundColor Cyan

foreach ($horario in $horarios) {
    $nomeTarefa = "Chamados_Extrator_$($horario.Replace(':',''))"

    Unregister-ScheduledTask -TaskName $nomeTarefa -Confirm:$false -ErrorAction SilentlyContinue

    $acao = New-ScheduledTaskAction `
        -Execute $python `
        -Argument "$script --auto" `
        -WorkingDirectory "C:\Users\gamaral\Desktop\Python\Chamados"

    $gatilho = New-ScheduledTaskTrigger `
        -Daily `
        -At $horario

    $config = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
        -StartWhenAvailable `
        -DontStopOnIdleEnd

    $principal = New-ScheduledTaskPrincipal `
        -UserId $env:USERNAME `
        -LogonType Interactive `
        -RunLevel Limited

    Register-ScheduledTask `
        -TaskName $nomeTarefa `
        -Action $acao `
        -Trigger $gatilho `
        -Settings $config `
        -Principal $principal `
        -Description "Extracao automatica Chamados MG Contecnica - $horario" `
        -Force

    Write-Host "Tarefa criada: $nomeTarefa ($horario)" -ForegroundColor Green
}

Write-Host ""
Write-Host "Todas as tarefas configuradas com sucesso!" -ForegroundColor Green
pause
