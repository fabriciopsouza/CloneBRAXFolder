import os
import shutil
import time
import datetime
import platform
import subprocess
from pathlib import Path
import logging
import re

class FileSync:
    def __init__(self):
        # Configurações de caminhos atualizados
        self.source_path = r"Y:\OPER_BAP_ESO\NP-1\01 - Carimbo AIVI (YSFECHAMENTO)\BRAX RPA"
        self.destination_path = r"\\172.31.9.186\interfaces\BW0924\OPER-ESO\1- SAP\YSMM_VI_HOMOLOG"
        self.file_pattern = "YSFECHAMENTO AIVI (ysvi_homolog_* oficial).xlsx"
        
        # Alvos para teste de conectividade
        self.ping_targets = [
            "172.31.9.186",
            "vibraenergia.com.br"
        ]

        # Configuração de logs
        self.user_home = str(Path.home())
        self.log_folder = os.path.join(self.user_home, "Documents", "CopyLatestExcelFile")
        self.log_file = os.path.join(self.log_folder, "CopyLatestExcelFile.log")
        
        # Criar pasta de logs se não existir
        os.makedirs(self.log_folder, exist_ok=True)
        
        # Configurar logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(message)s',
            handlers=[
                logging.FileHandler(self.log_file),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def ping(self, host):
        """Testa conectividade com ping"""
        param = '-n' if platform.system().lower() == 'windows' else '-c'
        command = ['ping', param, '1', host]
        try:
            return subprocess.call(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL) == 0
        except:
            return False

    def check_vpn_connection(self):
        """Verifica conexão VPN testando múltiplos alvos"""
        self.logger.info("Verificando conexão VPN...")
        
        # Testa ping para os alvos configurados
        for target in self.ping_targets:
            self.logger.info(f"Testando conexão com {target}...")
            if self.ping(target):
                self.logger.info(f"Conexão com {target} estabelecida!")
                return True
        
        # Testa acesso à pasta de origem
        try:
            if os.path.exists(self.source_path):
                self.logger.info("Acesso à pasta de origem confirmado!")
                return True
        except Exception as e:
            self.logger.error(f"Erro ao acessar pasta de origem: {e}")
        
        return False

    def get_year_from_filename(self, filename):
        """Extrai o ano do nome do arquivo usando regex"""
        try:
            # Procura por um padrão de 4 dígitos que represente um ano
            match = re.search(r'20\d{2}', filename)
            if match:
                return int(match.group())
            return None
        except:
            return None

    def find_latest_files(self):
        """Encontra os arquivos mais recentes de cada ano"""
        latest_files = {}  # {ano: (arquivo, data_modificacao)}
        
        try:
            self.logger.info(f"Buscando arquivos em: {self.source_path}")
            
            # Lista apenas arquivos do diretório principal (sem subpastas)
            files = [f for f in os.listdir(self.source_path) 
                    if os.path.isfile(os.path.join(self.source_path, f)) and 
                    "YSFECHAMENTO AIVI" in f and 
                    "oficial" in f and 
                    f.endswith('.xlsx')]

            self.logger.info(f"Total de arquivos encontrados: {len(files)}")

            for file in files:
                year = self.get_year_from_filename(file)
                if year:
                    file_path = os.path.join(self.source_path, file)
                    mod_time = os.path.getmtime(file_path)
                    
                    if year not in latest_files or mod_time > latest_files[year][1]:
                        latest_files[year] = (file, mod_time)
                        self.logger.info(f"Arquivo mais recente para {year}: {file}")
            
            # Ordena e retorna apenas os nomes dos arquivos mais recentes
            return [info[0] for year, info in sorted(latest_files.items())]
            
        except Exception as e:
            self.logger.error(f"Erro ao buscar arquivos: {e}")
            return []

    def copy_files(self):
        """Copia os arquivos mais recentes"""
        try:
            # Verifica se as pastas existem
            if not os.path.exists(self.source_path):
                self.logger.error(f"Pasta de origem não encontrada: {self.source_path}")
                return False

            if not os.path.exists(self.destination_path):
                self.logger.error(f"Pasta de destino não encontrada: {self.destination_path}")
                return False

            # Encontra os arquivos mais recentes
            latest_files = self.find_latest_files()
            
            self.logger.info(f"Encontrados {len(latest_files)} arquivos para copiar")

            # Copia cada arquivo
            for file in latest_files:
                try:
                    source = os.path.join(self.source_path, file)
                    destination = os.path.join(self.destination_path, file)
                    
                    # Verifica se o arquivo existe na origem
                    if not os.path.exists(source):
                        self.logger.error(f"Arquivo de origem não encontrado: {source}")
                        continue
                        
                    # Verifica se tem permissão para escrever no destino
                    if os.path.exists(destination):
                        try:
                            # Tenta remover o arquivo existente
                            os.remove(destination)
                            self.logger.info(f"Arquivo existente removido: {destination}")
                        except Exception as e:
                            self.logger.error(f"Erro ao remover arquivo existente {destination}: {e}")
                            continue
                    
                    # Copia o arquivo
                    shutil.copy2(source, destination)
                    self.logger.info(f"Arquivo copiado com sucesso: {file}")
                    
                    # Verifica se a cópia foi bem-sucedida
                    if os.path.exists(destination):
                        self.logger.info(f"Verificação: arquivo existe no destino: {destination}")
                    else:
                        self.logger.error(f"Verificação falhou: arquivo não existe no destino: {destination}")
                        
                except Exception as e:
                    self.logger.error(f"Erro ao copiar {file}: {e}")

            return True

        except Exception as e:
            self.logger.error(f"Erro durante a cópia de arquivos: {e}")
            return False

    def run(self):
        """Executa o processo completo"""
        self.logger.info("Iniciando processo de sincronização...")
        self.logger.info(f"Pasta de origem: {self.source_path}")
        self.logger.info(f"Pasta de destino: {self.destination_path}")

        # Verifica VPN com várias tentativas
        max_retries = 3
        for attempt in range(max_retries):
            if self.check_vpn_connection():
                break
            if attempt < max_retries - 1:
                self.logger.info(f"Tentativa {attempt + 1} de {max_retries} - Aguardando 30 segundos...")
                time.sleep(30)
            else:
                self.logger.error("Não foi possível estabelecer conexão VPN")
                return False

        # Copia os arquivos
        if self.copy_files():
            self.logger.info("Processo concluído com sucesso!")
        else:
            self.logger.error("Processo concluído com erros!")

        input("\nPressione Enter para sair...")

if __name__ == "__main__":
    sync = FileSync()
    sync.run()
