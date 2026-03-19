$origem = "C:\Users\ligomes\Downloads\painel_contabil_solucao_final\painel_contabil\dashboard_v3_8.html"
$destino = "C:\Users\ligomes\OneDrive - Globo Comunicação e Participações sa\Gestão de Eventos - Documentos\Gestão de Eventos_planejamento\Painel Contábil\2026\"

Set-Location "C:\Users\ligomes\Downloads\painel_contabil_solucao_final\painel_contabil"
python gerar_dashboard_v3.py

if (Test-Path $origem) {
    if (-not (Test-Path $destino)) {
        New-Item -ItemType Directory -Path $destino -Force
    }
    Copy-Item -Path $origem -Destination $destino -Force
    Write-Host "Arquivo copiado com sucesso."
} else {
    Write-Host "ERRO: Arquivo não gerado."
}