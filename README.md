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
