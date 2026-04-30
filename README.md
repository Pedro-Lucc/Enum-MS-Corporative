📌 ENUM MS – External User Enumeration (Microsoft 365)

Ferramenta em PowerShell para enumeração de contas corporativas do Microsoft 365 utilizando validação via endpoint público da Microsoft.

🧠 Sobre

O ENUM MS é um script que permite identificar usuários válidos em um domínio Microsoft 365 de forma externa (sem autenticação), analisando respostas do endpoint:

https://login.microsoftonline.com/common/GetCredentialType

--------------------------------------------------------------------------------------------------------

⚙️ Como funciona

A ferramenta segue o fluxo abaixo:

Recebe uma wordlist de usuários

Monta e-mails no formato:

usuario@dominio.com
Envia requisições para o endpoint da Microsoft
Analisa o campo IfExistsResult da resposta:
\n 
0 → Usuário EXISTE
1 → Usuário NÃO EXISTE
Outros → Desconhecido

--------------------------------------------------------------------------------------------------------

Uso
1. Criar wordlist

Exemplo (users.txt):

admin
joao
maria
suporte
financeiro

--------------------------------------------------------------------------------------------------------

2. Executar o script
   
Set-ExecutionPolicy Bypass -Scope Process -Force .\enum.ps1 -Domain empresa.com -UserList .\users.txt


Exibe resultados no console (com cores)
Salva os resultados em um arquivo .txt
Aplica delay aleatório (stealth)

--------------------------------------------------------------------------------------------------------

⚖️ Aviso Legal

Esta ferramenta deve ser utilizada apenas em ambientes autorizados.

O uso indevido pode violar leis locais e políticas de segurança.
