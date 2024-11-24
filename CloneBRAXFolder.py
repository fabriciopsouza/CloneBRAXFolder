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
        # Configurações de caminhos
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
        
        for target in self.ping_targets:
            self.logger.info(f"Testando conexão com {target}...")
            if self.ping(target):
                self.logger.info(f"Conexão com {target} estabelecida!")
                return True
        
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
            
            return [info[0] for year, info in sorted(latest_files.items())]
            
        except Exception as e:
            self.logger.error(f"Erro ao buscar arquivos: {e}")
            return []

    def copy_files(self):
        """Copia os arquivos mais recentes"""
        try:
            if not os.path.exists(self.source_path):
                self.logger.error(f"Pasta de origem não encontrada: {self.source_path}")
                return False

            if not os.path.exists(self.destination_path):
                self.logger.error(f"Pasta de destino não encontrada: {self.destination_path}")
                return False

            latest_files = self.find_latest_files()
            self.logger.info(f"Encontrados {len(latest_files)} arquivos para copiar")

            for file in latest_files:
                try:
                    source = os.path.join(self.source_path, file)
                    destination = os.path.join(self.destination_path, file)
                    
                    if not os.path.exists(source):
                        self.logger.error(f"Arquivo de origem não encontrado: {source}")
                        continue
                        
                    if os.path.exists(destination):
                        try:
                            os.remove(destination)
                            self.logger.info(f"Arquivo existente removido: {destination}")
                        except Exception as e:
                            self.logger.error(f"Erro ao remover arquivo existente {destination}: {e}")
                            continue
                    
                    shutil.copy2(source, destination)
                    self.logger.info(f"Arquivo copiado com sucesso: {file}")
                    
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
        """Executa o processo uma vez"""
        self.logger.info("Iniciando processo de sincronização...")
        self.logger.info(f"Pasta de origem: {self.source_path}")
        self.logger.info(f"Pasta de destino: {self.destination_path}")

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

        if self.copy_files():
            self.logger.info("Processo concluído com sucesso!")
        else:
            self.logger.error("Processo concluído com erros!")

    def should_run(self):
        """Verifica se é hora de executar baseado no horário"""
        now = datetime.datetime.now()
        target_hour = 9  # Executar às 9 da manhã
        target_minute = 0
        
        if (now.hour > target_hour) or (now.hour == target_hour and now.minute > target_minute):
            next_run = now.replace(day=now.day + 1, hour=target_hour, minute=target_minute, second=0, microsecond=0)
        else:
            next_run = now.replace(hour=target_hour, minute=target_minute, second=0, microsecond=0)
        
        wait_seconds = (next_run - now).total_seconds()
        
        if wait_seconds > 0:
            self.logger.info(f"Aguardando até {next_run.strftime('%H:%M')} para executar")
            time.sleep(wait_seconds)
        
        return True

    def run_scheduled(self):
        """Executa o processo continuamente com agendamento"""
        self.logger.info("Iniciando serviço de sincronização...")
        
        self.logger.info("Aguardando 60 segundos para inicialização completa do sistema...")
        time.sleep(60)
        
        while True:
            try:
                if self.should_run():
                    self.run()
                
                # Aguarda 1 hora antes de verificar novamente
                time.sleep(3600)
            except Exception as e:
                self.logger.error(f"Erro durante execução: {e}")
                time.sleep(300)  # Aguarda 5 minutos em caso de erro

if __name__ == "__main__":
    sync = FileSync()
    sync.run_scheduled()
