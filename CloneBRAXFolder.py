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
    # ... [todo o código anterior permanece igual até o método run()] ...

    def should_run(self):
        """Verifica se é hora de executar baseado no horário"""
        now = datetime.datetime.now()
        target_hour = 9  # Executar às 9 da manhã
        target_minute = 0
        
        # Se já passou do horário hoje, executar amanhã
        if (now.hour > target_hour) or (now.hour == target_hour and now.minute > target_minute):
            next_run = now.replace(day=now.day + 1, hour=target_hour, minute=target_minute, second=0, microsecond=0)
        else:
            next_run = now.replace(hour=target_hour, minute=target_minute, second=0, microsecond=0)
        
        # Calcula tempo até próxima execução
        wait_seconds = (next_run - now).total_seconds()
        
        if wait_seconds > 0:
            self.logger.info(f"Aguardando até {next_run.strftime('%H:%M')} para executar")
            time.sleep(wait_seconds)
        
        return True

    def run_scheduled(self):
        """Executa o processo continuamente com agendamento"""
        self.logger.info("Iniciando serviço de sincronização...")
        
        # Aguarda 60 segundos para garantir que a rede/VPN esteja disponível após o login
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
