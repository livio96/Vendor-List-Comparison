from ftplib import FTP, all_errors
import pandas as pd
import logging as log

from datetime import datetime
from vendor import Vendor


class Config:
    VENDOR = 'vendor'
    STATUS = 'status'
    CUSTOM_HEADER = 'custom header'
    LOOKUP = 'lookup'
    CHANGES = 'changes'
    RESULTS = 'results'
    LOCAL = 'local'
    FTP = 'ftp'
    OLD_EXCEL_FILE_EXT = 'old file ext'
    NEW_EXCEL_FILE_EXT = 'new file ext'
    OLD_EXCEL_FILE_SHEET_NAME = 'old file sheet'
    NEW_EXCEL_FILE_SHEET_NAME = 'new file sheet'
    EXTERNAL_ID_POSTFIX = 'postfix'

    SOURCE = 'source'
    SOURCE_FTP_URL = 'source ftp url'
    SOURCE_FTP_USER = 'source ftp user'
    SOURCE_FTP_PASS = 'source ftp pass'
    SOURCE_FTP_PORT = 'source ftp port'
    SOURCE_FTP_PATH = 'source ftp path'
    SOURCE_FTP_FILENAME = 'source ftp filename'

    RESULTS_FTP_URL = 'results ftp url'
    RESULTS_FTP_USER = 'results ftp user'
    RESULTS_FTP_PASS = 'results ftp pass'
    RESULTS_FTP_PORT = 'results ftp port'
    RESULTS_FTP_PATH = 'results ftp path'
    LOG_FILE_NAME = './Logs/' + datetime.now().strftime('Logs_%Y_%m_%d_%H_%M_%S.log')

    def __init__(self, name, sheet_name):
        self.__name = name
        self.__sheet_name = sheet_name
        self.__config_data_frame = None
        self.__columns = []
        self.__vendors = []
        self.__vendors = []
        self.__config_file_downloaded = False
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
        if not self.__config_file_downloaded:
            log.info(f'Start.....')
        try:
            # For XLXS Config File
            # self.__config_data_frame = pd.read_excel(self.__name, self.__sheet_name)

            # For CSV Config File
            self.__config_data_frame = pd.read_csv(self.__name, engine='python', encoding='unicode_escape')
            i = 1
            log.info(f'Reading config file sheet {self.__sheet_name} of {self.__name}')
            for column in self.__config_data_frame.columns:
                if i == 1 and column.lower() == Config.VENDOR:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 2 and column.lower() == Config.STATUS:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 3 and column.lower() == Config.CUSTOM_HEADER:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 4 and column.lower() == Config.LOOKUP:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 5 and column.lower() == Config.CHANGES:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 6 and column.lower() == Config.RESULTS:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 7 and column.lower() == Config.OLD_EXCEL_FILE_EXT:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 8 and column.lower() == Config.NEW_EXCEL_FILE_EXT:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 9 and column.lower() == Config.OLD_EXCEL_FILE_SHEET_NAME:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 10 and column.lower() == Config.NEW_EXCEL_FILE_SHEET_NAME:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 11 and column.lower() == Config.EXTERNAL_ID_POSTFIX:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 12 and column.lower() == Config.SOURCE:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 13 and column.lower() == Config.SOURCE_FTP_URL:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 14 and column.lower() == Config.SOURCE_FTP_USER:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 15 and column.lower() == Config.SOURCE_FTP_PASS:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 16 and column.lower() == Config.SOURCE_FTP_PORT:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 17 and column.lower() == Config.SOURCE_FTP_PATH:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 18 and column.lower() == Config.SOURCE_FTP_FILENAME:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 19 and column.lower() == Config.RESULTS_FTP_URL:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 20 and column.lower() == Config.RESULTS_FTP_USER:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 21 and column.lower() == Config.RESULTS_FTP_PASS:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 22 and column.lower() == Config.RESULTS_FTP_PORT:
                    self.__columns.append(column)
                    i += 1
                    continue
                elif i == 23 and column.lower() == Config.RESULTS_FTP_PATH:
                    self.__columns.append(column)
                    i += 1
                    continue
                else:
                    log.error(f'{self.__name} columns order should be Vendor, Status, Lookup, Changes, Old File Ext, New File Ext, Old File Sheet, New File Sheet, Postfix')
            for index, row in self.__config_data_frame.iterrows():
                vendor = Vendor()
                vendor.set_name(row[Config.VENDOR.title()])
                vendor.set_is_active(row[Config.STATUS.title()])
                vendor.set_custom_header(row[Config.CUSTOM_HEADER.title()])
                vendor.set_look_up(row[Config.LOOKUP.title()])
                vendor.set_changes(row[Config.CHANGES.title()])
                vendor.set_results(row[Config.RESULTS.title()])
                vendor.set_old_excel_file_extension(row[Config.OLD_EXCEL_FILE_EXT.title()])
                vendor.set_new_excel_file_extension(row[Config.NEW_EXCEL_FILE_EXT.title()])
                vendor.set_old_excel_file_sheet_name(row[Config.OLD_EXCEL_FILE_SHEET_NAME.title()])
                vendor.set_new_excel_file_sheet_name(row[Config.NEW_EXCEL_FILE_SHEET_NAME.title()])
                if row[Config.EXTERNAL_ID_POSTFIX.title()] != '' and str(row[Config.EXTERNAL_ID_POSTFIX.title()]) != 'nan':
                    vendor.set_external_id_postfix(row[Config.EXTERNAL_ID_POSTFIX.title()])

                vendor.set_source(row[Config.SOURCE.title()])
                vendor.set_source_ftp_url(row[Config.SOURCE_FTP_URL.title()])
                vendor.set_source_ftp_user(row[Config.SOURCE_FTP_USER.title()])
                vendor.set_source_ftp_pass(row[Config.SOURCE_FTP_PASS.title()])
                vendor.set_source_ftp_port(row[Config.SOURCE_FTP_PORT.title()])
                vendor.set_source_ftp_path(row[Config.SOURCE_FTP_PATH.title()])
                vendor.set_source_ftp_filename(row[Config.SOURCE_FTP_FILENAME.title()])

                vendor.set_results_ftp_url(row[Config.RESULTS_FTP_URL.title()])
                vendor.set_results_ftp_user(row[Config.RESULTS_FTP_USER.title()])
                vendor.set_results_ftp_pass(row[Config.RESULTS_FTP_PASS.title()])
                vendor.set_results_ftp_port(row[Config.RESULTS_FTP_PORT.title()])
                vendor.set_results_ftp_path(row[Config.RESULTS_FTP_PATH.title()])
                self.__vendors.append(vendor)

            vendors_str = ''.join(str(ven) for ven in self.__vendors)

            log.info(f'{len(self.__vendors)} row(s) found in {self.__name}, [{vendors_str}]')
        except FileNotFoundError:
            if not self.__config_file_downloaded:
                self.downloadConfigFileFromFTP()
            else:
                log.error(f'FileNotFoundError error occurred while reading config file {self.__name}')
        except ValueError as v:
            print(v)
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

                if vendor.get_source().lower() == Config.FTP:
                    vendor.download_source_excel_file_from_ftp_server(log)
                    if vendor.get_new_ftp_file_exists():
                        vendor.read_old_excel_file(log)
                        vendor.read_new_excel_file(log)
                        vendor.compare_data_frames(log)
                    else:
                        log.info(f'Skipping comparing files process, because FTP file is not downloaded.')
                elif vendor.get_source().lower() == Config.LOCAL:
                    vendor.read_old_excel_file(log)
                    vendor.read_new_excel_file(log)
                    vendor.compare_data_frames(log)
            else:
                log.info(f'Skip processing the {vendor.get_name()} having lookup column {vendor.get_look_up()} and '
                         f'columns to be checked is/are {vendor.get_changes()}')
        log.info(f'End.....')

    def downloadConfigFileFromFTP(self):
        host = 'telquestftp.com'
        username = 'admin@telquestftp.com'
        password = 'Shopping2016#'
        filename = 'config.csv'
        local_path = f'./{filename}'
        log.info(f'Downloading the config file {filename} from {host} FTP.')
        try:
            ftp = FTP(host)
            ftp.login(username, passwd=password)
            ftp.cwd('/')
            with open(local_path, "wb") as file:
                # use FTP's RETR command to download the file
                ftp.retrbinary(f'RETR {filename}', file.write)
                log.info(f'Config file {filename} download from {host} FTP and saved locally.')
            # file = open(self.__new_excel_file_path, 'wb')  # file to send
            # ftp.retrbinary('RETR ' + filename, file.write)  # send the file
            self.__config_file_downloaded = True
            ftp.quit()
            self.read_config_file()
        except all_errors as err:
            self.__config_file_downloaded = False
            log.error(f'Unable to download the config file {filename} from {host} FTP.')