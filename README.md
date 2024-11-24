Instruções de Instalação - Clone de Arquivos BRAX
O que você precisa fazer:

Instalar o Python:

Acesse: https://www.python.org/downloads/
Clique no botão amarelo grande "Download Python 3.12.x"
IMPORTANTE: Na tela de instalação, marque a caixa "Add Python to PATH"
Clique "Install Now"
Aguarde a instalação terminar


Baixar os arquivos necessários:

Salve este arquivo como CloneBRAXFolder.py na sua pasta Downloads
Salve este arquivo como agendar_tarefa.bat na sua pasta Downloads


Configurar o agendamento:

Dê duplo clique no arquivo agendar_tarefa.bat
Se aparecer uma mensagem do Windows Defender, clique em "Mais informações" e depois em "Executar assim mesmo"
Uma janela preta irá abrir e fechar rapidamente


Pronto! O programa irá:

Executar todos os dias às 9h da manhã
Copiar os arquivos mais recentes de cada ano
Criar logs em Documentos/CopyLatestExcelFile



Para executar manualmente:

Dê duplo clique no arquivo CloneBRAXFolder.py na pasta Downloads

Se precisar parar o agendamento:

Aperte Windows + R
Digite: taskschd.msc
Procure por "CopyExcelFiles_Daily"
Clique com botão direito -> Desabilitar

Em caso de problemas:

Verifique se marcou "Add Python to PATH" durante a instalação do Python
Certifique-se que está conectado à VPN
Os logs de erro ficam em: Documentos/CopyLatestExcelFile/CopyLatestExcelFile.log

Caso continue com problemas, entre em contato com o suporte técnico.
