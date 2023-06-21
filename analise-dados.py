import logging

# Configuração básica do logger
logging.basicConfig(filename=r'C:\temp\logfile.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Exemplo de registro de mensagens
logging.debug('Esta é uma mensagem de depuração')
logging.info('Esta é uma mensagem informativa')
logging.warning('Esta é uma mensagem de aviso')
logging.error('Esta é uma mensagem de erro')
logging.critical('Esta é uma mensagem crítica')

