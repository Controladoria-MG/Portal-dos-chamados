# ============================================================
# Agendador de Tarefas - MG Contecnica Chamados (Extrator + Base + Push)
# ============================================================

$python = (Get-Command python).Source
$script = "C:\Users\gamaral\Desktop\Python\Portal-dos-chamados\backend\extrato_chamados.py"
$workDir = "C:\Users\gamaral\Desktop\Python\Portal-dos-chamados"
$horarios = 9..18 | ForEach-Object { "{0:D2}:00" -f $_ }

Write-Host "Configurando tarefas agendadas..." -ForegroundColor Cyan

foreach ($horario in $horarios) {
    $nomeTarefa = "Chamados_Extrator_$($horario.Replace(':',''))"

    Unregister-ScheduledTask -TaskName $nomeTarefa -Confirm:$false -ErrorAction SilentlyContinue

    $acao = New-ScheduledTaskAction `
        -Execute $python `
        -Argument "$script --auto" `
        -WorkingDirectory $workDir

    $gatilho = New-ScheduledTaskTrigger `
        -Daily `
        -At $horario

    # Limite de 45 min: a rodada "--auto" faz 4 extracoes sequenciais + base +
    # push, e cada download pode levar ate 10 min se travar. Fica com folga
    # antes do proximo disparo (1h depois). IgnoreNew evita empilhar uma nova
    # instancia (e novos processos de Chrome) se a rodada anterior atrasar.
    $config = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 45) `
        -StartWhenAvailable `
        -DontStopOnIdleEnd `
        -MultipleInstances IgnoreNew

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
