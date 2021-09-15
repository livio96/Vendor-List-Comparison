from config import Config

CONFIG_FILE_NAME = 'config.xlsx'
CONFIG_FILE_SHEET_NAME = 'Sheet1'
config = Config(CONFIG_FILE_NAME, CONFIG_FILE_SHEET_NAME)
config.read_config_file()
config.process_vendors()
