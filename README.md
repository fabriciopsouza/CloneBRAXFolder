Guia de Uso do Script de Cópia Automática de Arquivos Excel
Este guia fornece instruções detalhadas, simples e concisas para que qualquer usuário não técnico possa executar corretamente o script em PowerShell que automatiza a cópia de arquivos Excel entre pastas específicas.

Sumário
Introdução
Pré-requisitos
Instruções de Instalação e Uso
Suporte
Licença
Introdução
Este script automatiza a cópia de arquivos Excel de uma pasta de origem para uma pasta de destino, verificando se há atualizações e agendando execuções diárias. Ele também garante que a cópia só ocorra quando a conexão VPN estiver ativa.

Pré-requisitos
Sistema Operacional: Windows 7 ou superior.
PowerShell: Versão 5.0 ou superior (já incluído no Windows 10).
Acesso às Pastas: Permissões de leitura na pasta de origem e de escrita na pasta de destino.
Conexão VPN: FortiClient VPN instalada e configurada.
Permissões Administrativas: Necessárias para agendar tarefas no Windows Task Scheduler.
Instruções de Instalação e Uso
Passo 1: Baixar o Script
Receba o arquivo do script CopyLatestExcelFile.ps1 fornecido pelo administrador ou suporte técnico.
Salve o arquivo em um local de fácil acesso, como a área de trabalho.
Passo 2: Verificar o PowerShell
Abrir o PowerShell:
Pressione Win + R, digite powershell e pressione Enter.
Verificar a Versão:
No prompt do PowerShell, digite $PSVersionTable.PSVersion e pressione Enter.
Certifique-se de que a versão principal (Major) seja 5 ou superior.
Passo 3: Ajustar a Política de Execução
Alterar a Política Temporariamente:
No PowerShell, digite Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass e pressione Enter.
Quando solicitado, digite Y e pressione Enter.
Passo 4: Executar o Script
Conectar à VPN:
Certifique-se de que o FortiClient VPN esteja conectado.
Executar o Script:
Navegue até o local onde o script foi salvo.
Clique com o botão direito no arquivo CopyLatestExcelFile.ps1 e selecione Executar com PowerShell.
Aguardar a Execução:
O script será executado em segundo plano. Ele pode levar alguns minutos para concluir a primeira execução.
Passo 5: Verificar o Log
Abrir o Log:
Navegue até a pasta Meus Documentos.
Abra o arquivo CopyLatestExcelFile.log para verificar as ações realizadas pelo script.
Passo 6: Confirmar o Agendamento da Tarefa
Abrir o Agendador de Tarefas:
Pressione Win + R, digite taskschd.msc e pressione Enter.
Localizar a Tarefa:
Na biblioteca do Agendador de Tarefas, procure pela tarefa CopyLatestExcelFile.
Verificar os Detalhes:
Certifique-se de que a tarefa está agendada para executar diariamente no mesmo horário.
Suporte
Se você encontrar problemas ou tiver dúvidas, entre em contato com o suporte técnico:

E-mail: suporte@empresa.com
Telefone: (11) 1234-5678
Licença
Este script é fornecido "no estado em que se encontra", sem garantias de qualquer tipo, expressas ou implícitas. O uso é permitido apenas para fins internos da empresa. Consulte a seção Licença para obter mais detalhes.

========================================================

Automatic Excel File Copy Script User Guide
This guide provides detailed, simple, and concise instructions for any non-technical user to correctly execute the PowerShell script that automates copying Excel files between specific folders.

Table of Contents
Introduction
Prerequisites
Installation and Usage Instructions
Support
License
Introduction
This script automates the copying of Excel files from a source folder to a destination folder, checking for updates and scheduling daily executions. It also ensures that the copy only occurs when the VPN connection is active.

Prerequisites
Operating System: Windows 7 or later.
PowerShell: Version 5.0 or higher (included in Windows 10).
Folder Access: Read permissions on the source folder and write permissions on the destination folder.
VPN Connection: FortiClient VPN installed and configured.
Administrative Permissions: Required to schedule tasks in Windows Task Scheduler.
Installation and Usage Instructions
Step 1: Download the Script
Obtain the CopyLatestExcelFile.ps1 script file provided by your administrator or technical support.
Save the file in an easily accessible location, such as the desktop.
Step 2: Verify PowerShell
Open PowerShell:
Press Win + R, type powershell, and press Enter.
Check Version:
In the PowerShell prompt, type $PSVersionTable.PSVersion and press Enter.
Ensure the major version (Major) is 5 or higher.
Step 3: Adjust Execution Policy
Temporarily Change Policy:
In PowerShell, type Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass and press Enter.
When prompted, type Y and press Enter.
Step 4: Run the Script
Connect to VPN:
Ensure that the FortiClient VPN is connected.
Execute the Script:
Navigate to where the script is saved.
Right-click on CopyLatestExcelFile.ps1 and select Run with PowerShell.
Wait for Execution:
The script will run in the background. It may take a few minutes to complete the first execution.
Step 5: Check the Log
Open the Log:
Navigate to your My Documents folder.
Open the file CopyLatestExcelFile.log to review the actions performed by the script.
Step 6: Confirm Task Scheduling
Open Task Scheduler:
Press Win + R, type taskschd.msc, and press Enter.
Locate the Task:
In the Task Scheduler Library, look for the task named CopyLatestExcelFile.
Verify Details:
Ensure the task is scheduled to run daily at the same time.
Support
If you encounter issues or have questions, please contact technical support:

Email: support@company.com
Phone: (123) 456-7890
License
This script is provided "as is," without warranty of any kind, express or implied. Usage is permitted for internal company purposes only. Refer to the License section for more details.
