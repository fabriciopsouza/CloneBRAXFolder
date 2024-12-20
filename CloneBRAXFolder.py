import os
import shutil
import time
import platform
import subprocess
from pathlib import Path
import logging
import re
import json
from datetime import datetime

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

        # Configuração de logs e controle
        self.user_home = str(Path.home())
        self.log_folder = os.path.join(self.user_home, "Documents", "CopyLatestExcelFile")
        self.log_file = os.path.join(self.log_folder, "CopyLatestExcelFile.log")
        self.control_file = os.path.join(self.log_folder, "last_copied_files.json")
        
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

    def load_control_file(self):
        """Carrega informações do último arquivo copiado"""
        try:
            if os.path.exists(self.control_file):
                with open(self.control_file, 'r') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            self.logger.error(f"Erro ao ler arquivo de controle: {e}")
            return {}

    def save_control_file(self, control_data):
        """Salva informações do arquivo copiado"""
        try:
            with open(self.control_file, 'w') as f:
                json.dump(control_data, f, indent=2)
        except Exception as e:
            self.logger.error(f"Erro ao salvar arquivo de controle: {e}")

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
        """Encontra os arquivos mais recentes de cada ano que precisam ser atualizados"""
        latest_files = {}  # {ano: (arquivo, data_modificacao)}
        control_data = self.load_control_file()
        files_to_update = []
        
        try:
            self.logger.info(f"Buscando arquivos em: {self.source_path}")
            
            files = [f for f in os.listdir(self.source_path) 
                    if os.path.isfile(os.path.join(self.source_path, f)) and 
                    "YSFECHAMENTO AIVI" in f and 
                    "oficial" in f and 
                    f.endswith('.xlsx')]

            self.logger.info(f"Total de arquivos encontrados: {len(files)}")

            # Primeiro encontra o mais recente de cada ano
            for file in files:
                year = self.get_year_from_filename(file)
                if year:
                    file_path = os.path.join(self.source_path, file)
                    mod_time = os.path.getmtime(file_path)
                    
                    if year not in latest_files or mod_time > latest_files[year][1]:
                        latest_files[year] = (file, mod_time)

            # Depois verifica quais precisam ser atualizados
            for year, (file, mod_time) in latest_files.items():
                year_str = str(year)
                last_copied = control_data.get(year_str, {})
                last_mod_time = last_copied.get('mod_time', 0)
                
                if mod_time > last_mod_time:
                    files_to_update.append(file)
                    # Atualiza o controle com o novo timestamp
                    control_data[year_str] = {
                        'filename': file,
                        'mod_time': mod_time,
                        'last_copied': datetime.now().isoformat()
                    }
                    self.logger.info(f"Arquivo mais recente para {year} precisa ser atualizado: {file}")
                else:
                    self.logger.info(f"Arquivo de {year} não precisa ser atualizado")
            
            # Salva as atualizações no arquivo de controle
            self.save_control_file(control_data)
            
            return files_to_update
            
        except Exception as e:
            self.logger.error(f"Erro ao buscar arquivos: {e}")
            return []

    def copy_files(self):
        """Copia apenas os arquivos que precisam ser atualizados"""
        try:
            if not os.path.exists(self.source_path):
                self.logger.error(f"Pasta de origem não encontrada: {self.source_path}")
                return False

            if not os.path.exists(self.destination_path):
                self.logger.error(f"Pasta de destino não encontrada: {self.destination_path}")
                return False

            files_to_update = self.find_latest_files()
            
            if not files_to_update:
                self.logger.info("Nenhum arquivo precisa ser atualizado")
                return True

            self.logger.info(f"Encontrados {len(files_to_update)} arquivos para atualizar")

            for file in files_to_update:
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

        # Tenta conectar à VPN algumas vezes
        max_retries = 3
        for attempt in range(max_retries):
            if self.check_vpn_connection():
                break
            if attempt < max_retries - 1:
                self.logger.info(f"Tentativa {attempt + 1} de {max_retries} - Aguardando 30 segundos...")
                time.sleep(30)
            else:
                self.logger.error("Não foi possível estabelecer conexão VPN")
                input("\nPressione Enter para sair...")
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
