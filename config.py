import pandas as pd
import logging as log

from datetime import datetime
from vendor import Vendor


class Config:
    VENDOR = 'vendor'
    STATUS = 'status'
    LOOKUP = 'lookup'
    CHANGES = 'changes'
    OLD_EXCEL_FILE_EXT = 'old file ext'
    NEW_EXCEL_FILE_EXT = 'new file ext'
    OLD_EXCEL_FILE_SHEET_NAME = 'old file sheet'
    NEW_EXCEL_FILE_SHEET_NAME = 'new file sheet'
    EXTERNAL_ID_POSTFIX = 'postfix'

    RESULTS_FTP_URL  = 'results ftp url'
    RESULTS_FTP_USER = 'results ftp user'
    RESULTS_FTP_PASS = 'results ftp pass'
    RESULTS_FTP_PORT = 'results ftp port'

    LOG_FILE_NAME = './Logs/' + datetime.now().strftime('Logs_%Y_%m_%d_%H_%M_%S.log')

    def __init__(self, name, sheet_name):
        self.__name = name
        self.__sheet_name = sheet_name
        self.__config_data_frame = None
        self.__columns = []
        self.__vendors = []
        log.basicConfig(filename=Config.LOG_FILE_NAME, format='%(asctime)s : %(message)s', filemode='w', level=log.INFO)

    def __str__(self):
        return f'({self.__name}, {self.__sheet_name})'

    def get_name(self):
        return self.__name

    def set_name(self, name):
        self.__name = name

    def get_sheet_name(self):
        return self.__sheet_name

    def set_sheet_name(self, sheet_name):
        self.__sheet_name = sheet_name

    def read_config_file(self):
        log.info(f'Start.....')
        log.info(f'Reading config file sheet {self.__sheet_name} of {self.__name}')
        try:
            self.__config_data_frame = pd.read_excel(self.__name, self.__sheet_name)
            i = 1
            for column in self.__config_data_frame.columns:
                if i == 1 and column.lower() == Config.VENDOR:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 2 and column.lower() == Config.STATUS:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 3 and column.lower() == Config.LOOKUP:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 4 and column.lower() == Config.CHANGES:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 5 and column.lower() == Config.OLD_EXCEL_FILE_EXT:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 6 and column.lower() == Config.NEW_EXCEL_FILE_EXT:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 7 and column.lower() == Config.OLD_EXCEL_FILE_SHEET_NAME:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 8 and column.lower() == Config.NEW_EXCEL_FILE_SHEET_NAME:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 9 and column.lower() == Config.EXTERNAL_ID_POSTFIX:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 10 and column.lower() == Config.RESULTS_FTP_URL:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 11 and column.lower() == Config.RESULTS_FTP_USER:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 12 and column.lower() == Config.RESULTS_FTP_PASS:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 13 and column.lower() == Config.RESULTS_FTP_PORT:
                    self.__columns.append(column)
                    i += 1
                    continue
                else:
                    log.error(f'{self.__name} columns order should be Vendor, Status, Lookup, Changes, Old File Ext, New File Ext, Old File Sheet, New File Sheet, Postfix')

            for index, row in self.__config_data_frame.iterrows():
                vendor = Vendor()
                vendor.set_name(row[Config.VENDOR.title()])
                vendor.set_is_active(row[Config.STATUS.title()])
                vendor.set_look_up(row[Config.LOOKUP.title()])
                vendor.set_changes(row[Config.CHANGES.title()])
                vendor.set_old_excel_file_extension(row[Config.OLD_EXCEL_FILE_EXT.title()])
                vendor.set_new_excel_file_extension(row[Config.NEW_EXCEL_FILE_EXT.title()])
                vendor.set_old_excel_file_sheet_name(row[Config.OLD_EXCEL_FILE_SHEET_NAME.title()])
                vendor.set_new_excel_file_sheet_name(row[Config.NEW_EXCEL_FILE_SHEET_NAME.title()])
                if row[Config.EXTERNAL_ID_POSTFIX.title()] != '' and str(row[Config.EXTERNAL_ID_POSTFIX.title()]) != 'nan':
                    vendor.set_external_id_postfix(row[Config.EXTERNAL_ID_POSTFIX.title()])

                vendor.set_results_ftp_url(row[Config.RESULTS_FTP_URL.title()])
                vendor.set_results_ftp_user(row[Config.RESULTS_FTP_USER.title()])
                vendor.set_results_ftp_pass(row[Config.RESULTS_FTP_PASS.title()])
                vendor.set_results_ftp_port(row[Config.RESULTS_FTP_PORT.title()])
                self.__vendors.append(vendor)

            vendors_str = ''.join(str(ven) for ven in self.__vendors)

            log.info(f'{len(self.__vendors)} row(s) found in {self.__name}, [{vendors_str}]')
        except FileNotFoundError:
            log.error(f'FileNotFoundError error occurred while reading config file {self.__name}')
        except ValueError:
            log.error(f'ValueError error occurred while reading config file sheet {self.__sheet_name}')
        except AttributeError:
            log.error(
                f'AttributeError error occurred while reading config file sheet {self.__sheet_name} of {self.__name}')
        except ImportError:
            log.error(f'Please install missing Python Libraries.')

    def process_vendors(self):
        for vendor in self.__vendors:
            if vendor.get_is_active():
                log.info(f'Processing the {vendor.get_name()} having lookup column {vendor.get_look_up()} and columns '
                         f'to be checked is/are {vendor.get_changes()}')
                vendor.get_old_excel_file_path()
                vendor.get_new_excel_file_path()
                vendor.get_results_csv_file_path()
                vendor.read_old_excel_file(log)
                vendor.read_new_excel_file(log)
                vendor.compare_data_frames(log)
            else:
                log.info(f'Skip processing the {vendor.get_name()} having lookup column {vendor.get_look_up()} and '
                         f'columns to be checked is/are {vendor.get_changes()}')
        log.info(f'End.....')
