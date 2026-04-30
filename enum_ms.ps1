# ============================================

# EXTERNAL USER ENUM (M365) - FINAL VERSION

# ============================================

param(
[string]$Domain = "empresa.com",
[string]$UserList = ".\users.txt"
)

# =========================

# PREPARAÇÃO

# =========================

if (!(Test-Path $UserList)) {
Write-Host "[ERRO] Wordlist não encontrada: $UserList" -ForegroundColor Red
exit
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = "enum_external_$timestamp.txt"

Write-Host "==========================================" -ForegroundColor DarkGray
Write-Host "   ███████╗███╗   ██╗██╗   ██╗███╗   ███╗" -ForegroundColor Cyan
Write-Host "   ██╔════╝████╗  ██║██║   ██║████╗ ████║" -ForegroundColor Cyan
Write-Host "   █████╗  ██╔██╗ ██║██║   ██║██╔████╔██║" -ForegroundColor Cyan
Write-Host "   ██╔══╝  ██║╚██╗██║██║   ██║██║╚██╔╝██║" -ForegroundColor Cyan
Write-Host "   ███████╗██║ ╚████║╚██████╔╝██║ ╚═╝ ██║" -ForegroundColor Cyan
Write-Host "   ╚══════╝╚═╝  ╚═══╝ ╚═════╝ ╚═╝     ╚═╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "             ENUM MS" -ForegroundColor Yellow
Write-Host "------------------------------------------" -ForegroundColor DarkGray
Write-Host " ferramenta para enumeração de contas corporativas da Microsoft" -ForegroundColor Gray
Write-Host "==========================================" -ForegroundColor DarkGray

Write-Host " - COMO FUNCIONA: " -ForegroundColor Yellow
Write-Host " - Recebe uma lista de usuários (wordlist)" -ForegroundColor Gray 
Write-Host " - Monta emails no padrão usuario@dominio" -ForegroundColor Gray
Write-Host " - Consulta endpoint de autenticação da Microsoft" -ForegroundColor Gray 
Write-Host " - Analisa a resposta para identificar contas válidas" -ForegroundColor Gray
Write-Host ""

Write-Host "[+] Iniciando enumeração externa..." -ForegroundColor Cyan
Write-Host "[+] Domínio: $Domain"
Write-Host "[+] Wordlist: $UserList"
Write-Host "[+] Output: $outputFile`n"

$url = "https://login.microsoftonline.com/common/GetCredentialType"

$valid = 0
$invalid = 0
$unknown = 0

# =========================

# ENUMERAÇÃO

# =========================

foreach ($user in Get-Content $UserList) {

if ([string]::IsNullOrWhiteSpace($user)) { continue }

$email = "$user@$Domain"

$body = @{
    Username = $email
} | ConvertTo-Json

try {
    $response = Invoke-RestMethod -Method POST -Uri $url -Body $body -ContentType "application/json"

    $status = "DESCONHECIDO"

    if ($response.IfExistsResult -eq 0) {
        $status = "VALIDO"
        $valid++
    }
    elseif ($response.IfExistsResult -eq 1) {
        $status = "INVALIDO"
        $invalid++
    }
    else {
        $unknown++
    }

    $line = "$email -> $status"

    # Console colorido
    if ($status -eq "VALIDO") {
        Write-Host $line -ForegroundColor Green
    }
    elseif ($status -eq "INVALIDO") {
        Write-Host $line -ForegroundColor DarkGray
    }
    else {
        Write-Host $line -ForegroundColor Yellow
    }

    # Salva no TXT
    Add-Content -Path $outputFile -Value $line

} catch {
    $line = "$email -> ERRO"
    Write-Host $line -ForegroundColor Red
    Add-Content -Path $outputFile -Value $line
}

# Delay aleatório (stealth)
Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 800)


}

# =========================

# RESUMO FINAL

# =========================

$summary = "=========================`n" +
           "RESUMO`n" +
"=========================`n" +
           "VALIDOS: $valid`n" +
"INVALIDOS: $invalid`n" +
           "DESCONHECIDOS: $unknown`n" +
"=========================`n"

Write-Host "`n$summary" -ForegroundColor Cyan
Add-Content -Path $outputFile -Value "`n$summary"

Write-Host "[FINALIZADO] Resultado salvo em $outputFile" -ForegroundColor Green
