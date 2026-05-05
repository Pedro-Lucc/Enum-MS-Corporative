# ENUM MS Corporate

Script PowerShell para enumeracao de contas corporativas Microsoft e, opcionalmente, teste controlado de login com geracao de relatorios em Excel.

> Use somente em ambientes proprios ou com autorizacao formal. O modo `-TestLogin` realiza tentativas reais de autenticacao e pode gerar alertas, bloqueios, throttling ou impacto em contas.

## Funcionalidades

- Enumera usuarios em um dominio Microsoft/Entra ID.
- Classifica usuarios como `VALIDO`, `INVALIDO` ou `DESCONHECIDO`.
- Detecta sinais de MFA quando presentes na resposta.
- Opcionalmente testa login para usuarios validos.
- Registra cada tentativa de senha em uma linha no Excel.
- Separa resultados em um arquivo xlsx.
- Suporta proxy unico ou lista de proxies com rotação.
- Possui ajuda integrada com `-h` ou `-Help`.

## Requisitos

- Windows PowerShell 5.1 ou PowerShell 7+.
- Acesso HTTPS para `https://login.microsoftonline.com`.
- Permissao para executar scripts PowerShell.
- Permissao para instalar modulos no escopo do usuario, caso ainda nao existam.

Modulos usados:

```powershell
MSAL.PS
ImportExcel
```

O script tenta instalar automaticamente os modulos ausentes:

```powershell
Install-Module MSAL.PS -Force -Scope CurrentUser -AllowClobber
Install-Module ImportExcel -Force -Scope CurrentUser -AllowClobber
```

Se a execucao de scripts estiver bloqueada:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

## Arquivos Necessarios

### `users.txt`

Lista de usuarios, um por linha, sem o dominio:

```text
jsilva
maria1
joseal
```

Com `-Domain empresa.onmicrosoft.com`, o script monta:

```text
jsilva@empresa.onmicrosoft.com
maria1@empresa.onmicrosoft.com
joseal@empresa.onmicrosoft.com
```

### `senhas.txt`

Necessario apenas quando usar `-TestLogin`.

```text
Senha123
P@ssw0rd!
OutraSenha
```

## Uso Rapido

Mostrar ajuda:

```powershell
.\enum_passs_brute_fixed_v2.ps1 -h
```

Enumerar usuarios:

```powershell
.\enum_passs_brute_fixed_v2.ps1 -Domain empresa.onmicrosoft.com -UserList .\users.txt
```

Enumerar e perguntar antes de testar login:

```powershell
.\enum_passs_brute_fixed_v2.ps1 -Domain empresa.onmicrosoft.com -UserList .\users.txt -TestLogin
```

Enumerar e testar login sem perguntar:

```powershell
.\enum_passs_brute_fixed_v2.ps1 -Domain empresa.onmicrosoft.com -UserList .\users.txt -TestLogin -nq
```

## Parametros

| Parametro | Tipo | Descricao |
| --- | --- | --- |
| `-Domain`                     | string    | Dominio alvo. Exemplo: `empresa.onmicrosoft.com`. |
| `-UserList`                   | string    | Caminho do arquivo com usuarios. Padrao: `.\users.txt`. |
| `-TestLogin`                  | switch    | Ativa o teste de login para usuarios validos. |
| `-NoQuestion`                 | switch    | Nao pergunta antes de testar login. |
| `-nq`                         | alias     | Alias de `-NoQuestion`. |
| `-Proxy`                      | string    | Usa um proxy unico. |
| `-ProxyList`                  | string[]  | Usa um ou mais proxies com rotacao. |
| `-ProxyUseDefaultCredentials` | switch    | Usa credenciais Windows no proxy. |
| `-Help`                       | switch    | Mostra ajuda. |
| `-h`                          | alias     | Alias de `-Help`. |

## Uso com Proxy

Proxy unico:

```powershell
.\enum_passs_brute_fixed_v2.ps1 -Domain empresa.onmicrosoft.com -UserList .\users.txt -TestLogin -Proxy "http://127.0.0.1:8080"
```

Lista de proxies com rotacao:

```powershell
.\enum_passs_brute_fixed_v2.ps1 -Domain empresa.onmicrosoft.com -UserList .\users.txt -TestLogin -ProxyList "http://127.0.0.1:8080","http://127.0.0.1:8081"
```

Proxy com credenciais integradas do Windows:

```powershell
.\enum_passs_brute_fixed_v2.ps1 -Domain empresa.onmicrosoft.com -UserList .\users.txt -TestLogin -Proxy "http://proxy.local:8080" -ProxyUseDefaultCredentials
```

Se o proxy fizer inspecao HTTPS, instale a CA do proxy como confiavel no Windows. Caso contrario, pode ocorrer erro de confiança SSL/TLS.

## Saidas Geradas

O script cria um arquivo geral no diretorio atual:

```text
enum_ms_YYYYMMDD_HHMMSS.xlsx
```

Tambem cria a pasta:

```text
enum_resultados\
```

Com arquivos separados:

```text
validos.xlsx
senha_incorreta.xlsx
mfa_requerido.xlsx
bloqueados.xlsx
outros.xlsx
resumo.xlsx
```

## Estrutura das Linhas no Excel

Cada tentativa de login e registrada em linha propria:

| DataHora | Email | Status | MFA | LoginStatus | Senha |
| --- | --- | --- | --- | --- | --- |
| 2026-05-05 16:30:00 | usuario@empresa.onmicrosoft.com | VALIDO | MFA_REQUERIDO | SENHA_INCORRETA | Senha123 |
| 2026-05-05 16:30:02 | usuario@empresa.onmicrosoft.com | VALIDO | MFA_REQUERIDO | LOGIN_OK MAS MFA REQUERIDO | P@ssw0rd! |

Quando `-TestLogin` nao e usado, ou nenhuma tentativa e registrada, a coluna `Senha` fica vazia.

## Status Possiveis

### `Status`

- `VALIDO`
- `INVALIDO`
- `DESCONHECIDO`
- `ERRO`

### `MFA`

- `NAO_DETECTADO`
- `MFA_REGISTRADO`
- `MFA_REQUERIDO`
- `MFA_DETECTADO`
- `MFA_POSSIVEL`

### `LoginStatus`

- `NAO_TESTADO`
- `SENHA_INCORRETA`
- `LOGIN_OK`
- `LOGIN_OK MAS MFA REQUERIDO`
- `CONTA_BLOQUEADA`
- `USUARIO_NAO_ENCONTRADO`
- `ERRO`

## Solucao de Problemas

### Falta permissao para executar o script

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### Modulo nao instala

Verifique a conectividade com a PowerShell Gallery e tente instalar manualmente:

```powershell
Install-Module MSAL.PS -Scope CurrentUser -Force
Install-Module ImportExcel -Scope CurrentUser -Force
```

### Erro SSL/TLS

Force TLS 1.2 na sessao atual:

```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
```

Se estiver usando proxy com interceptacao HTTPS, instale o certificado CA do proxy como confiavel.

### Proxy nao funciona

Teste o proxy isoladamente:

```powershell
Invoke-RestMethod -Uri "https://login.microsoftonline.com" -Proxy "http://127.0.0.1:8080"
```

## Boas Praticas

- Use somente com autorização.
- Evite listas grandes de senhas.
- Considere politicas de bloqueio de conta do ambiente.
- Documente escopo, janela de teste e responsáveis antes da execução.

## Licenca

Adicione a licenca apropriada antes de publicar o projeto no GitHub.
