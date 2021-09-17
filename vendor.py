from ftplib import FTP, all_errors
import pandas as pd
import os


class Vendor:
    OLD_FILE_NAME = ' - old'
    NEW_FILE_NAME = ' - new'

    OUTPUT_FILE_EXT = '.csv'

    CSV_FILE_EXT = '.csv'
    XLS_FILE_EXT = '.xls'
    XLSX_FILE_EXT = '.xlsx'

    RESULTS_FOLDER = 'Results'
    RESULTS_FILE_POST_FIX = ' - results'
    ACTIVE = 'Active'
    IN_ACTIVE = 'Inactive'

    def __init__(self):
        self.__name = None
        self.__is_active = False
        self.__look_up = None
        self.__changes = []

        self.__old_excel_file_path = None
        self.__new_excel_file_path = None
        self.__results_csv_file_path = None

        self.__old_excel_file_sheet_name = 'Sheet1'
        self.__new_excel_file_sheet_name = 'Sheet1'

        self.__old_excel_file_extension = '.xlsx'
        self.__new_excel_file_extension = '.xlsx'

        self.__old_excel_file_exists = False
        self.__new_excel_file_exists = False

        self.__new_ftp_file_exists = False

        self.__old_excel_file_data_frame = None
        self.__new_excel_file_data_frame = None
        self.__external_id_postfix = None

        self.__source_ftp_url = None
        self.__source_ftp_user = None
        self.__source_ftp_pass = None
        self.__source_ftp_port = None
        self.__source_ftp_path = None
        self.__source_ftp_filename = None

        self.__results_ftp_url = None
        self.__results_ftp_user = None
        self.__results_ftp_pass = None
        self.__results_ftp_port = None
        self.__results_ftp_path = None

        self.__different_rows_indices = []

    def __str__(self):
        return f'({self.__name}, {self.__is_active}, {self.__look_up}, {self.__changes})'

    def get_old_excel_file_sheet_name(self):
        return self.__old_excel_file_sheet_name

    def set_old_excel_file_sheet_name(self, old_excel_file_sheet_name):
        self.__old_excel_file_sheet_name = old_excel_file_sheet_name

    def get_new_excel_file_sheet_name(self):
        return self.__new_excel_file_sheet_name

    def set_new_excel_file_sheet_name(self, new_excel_file_sheet_name):
        self.__new_excel_file_sheet_name = new_excel_file_sheet_name

    def get_old_excel_file_extension(self):
        return self.__old_excel_file_extension

    def set_old_excel_file_extension(self, old_excel_file_extension):
        self.__old_excel_file_extension = old_excel_file_extension

    def get_new_excel_file_extension(self):
        return self.__new_excel_file_extension

    def set_new_excel_file_extension(self, new_excel_file_extension):
        self.__new_excel_file_extension = new_excel_file_extension

    def get_source_ftp_url(self):
        return self.__source_ftp_url

    def set_source_ftp_url(self, source_ftp_url):
        self.__source_ftp_url = source_ftp_url

    def get_source_ftp_user(self):
        return self.__source_ftp_user

    def set_source_ftp_user(self, source_ftp_user):
        self.__source_ftp_user = source_ftp_user

    def get_source_ftp_pass(self):
        return self.__source_ftp_pass

    def set_source_ftp_pass(self, source_ftp_pass):
        self.__source_ftp_pass = source_ftp_pass

    def get_source_ftp_port(self):
        return self.__source_ftp_port

    def set_source_ftp_port(self, source_ftp_port):
        self.__source_ftp_port = source_ftp_port

    def get_source_ftp_path(self):
        return self.__source_ftp_path

    def set_source_ftp_path(self, source_ftp_path):
        self.__source_ftp_path = source_ftp_path

    def get_source_ftp_filename(self):
        return self.__source_ftp_filename

    def set_source_ftp_filename(self, source_ftp_filename):
        self.__source_ftp_filename = source_ftp_filename

    def get_results_ftp_url(self):
        return self.__results_ftp_url

    def set_results_ftp_url(self, results_ftp_url):
        self.__results_ftp_url = results_ftp_url

    def get_results_ftp_user(self):
        return self.__results_ftp_user

    def set_results_ftp_user(self, results_ftp_user):
        self.__results_ftp_user = results_ftp_user

    def get_results_ftp_pass(self):
        return self.__results_ftp_pass

    def set_results_ftp_pass(self, results_ftp_pass):
        self.__results_ftp_pass = results_ftp_pass

    def get_results_ftp_port(self):
        return self.__results_ftp_port

    def set_results_ftp_port(self, results_ftp_port):
        self.__results_ftp_port = results_ftp_port

    def get_results_ftp_path(self):
        return self.__results_ftp_path

    def set_results_ftp_path(self, results_ftp_path):
        self.__results_ftp_path = results_ftp_path

    def get_new_ftp_file_exists(self):
        return self.__new_ftp_file_exists

    def set_new_ftp_file_exists(self, new_ftp_file_exists):
        self.__new_ftp_file_exists = new_ftp_file_exists

    def get_name(self):
        return self.__name

    def set_name(self, name):
        self.__name = name

    def get_is_active(self):
        return self.__is_active

    def set_is_active(self, is_active):
        if is_active == Vendor.ACTIVE:
            self.__is_active = True
        elif is_active == Vendor.IN_ACTIVE:
            self.__is_active = False

    def get_look_up(self):
        return self.__look_up

    def set_look_up(self, look_up):
        self.__look_up = look_up

    def get_changes(self):
        return self.__changes

    def set_changes(self, changes):
        if changes.find(',') > -1:
            self.__changes = changes.split(',')
        else:
            self.__changes.append(changes)

    def get_external_id_postfix(self):
        return self.__external_id_postfix

    def set_external_id_postfix(self, external_id_postfix):
        self.__external_id_postfix = external_id_postfix

    def get_old_excel_file_path(self):
        self.__old_excel_file_path = f'./{self.__name}/{self.__name}{Vendor.OLD_FILE_NAME}{self.__old_excel_file_extension}'
        print(self.__old_excel_file_path)
        return self.__old_excel_file_path

    def get_new_excel_file_path(self):
        self.__new_excel_file_path = f'./{self.__name}/{self.__name}{Vendor.NEW_FILE_NAME}{self.__new_excel_file_extension}'
        print(self.__new_excel_file_path)
        return self.__new_excel_file_path

    def get_results_csv_file_path(self):
        self.__results_csv_file_path = f'./{Vendor.RESULTS_FOLDER}/{self.__name}{Vendor.RESULTS_FILE_POST_FIX}{Vendor.OUTPUT_FILE_EXT} '
        return self.__results_csv_file_path

    def read_old_excel_file(self, log):
        log.info(
            f'Reading {self.__name} - old data file sheet {self.__old_excel_file_sheet_name} at {self.__old_excel_file_path}')
        try:
            if self.__old_excel_file_extension == Vendor.XLS_FILE_EXT or self.__old_excel_file_extension == Vendor.XLSX_FILE_EXT:
                self.__old_excel_file_data_frame = pd.read_excel(self.__old_excel_file_path,
                                                                 self.__old_excel_file_sheet_name)
                self.__old_excel_file_exists = True
            elif self.__old_excel_file_extension == Vendor.CSV_FILE_EXT:
                print("ELIF")
                print(self.__old_excel_file_path)
                self.__old_excel_file_data_frame = pd.read_csv(self.__old_excel_file_path)
                print(self.__old_excel_file_data_frame.shape)
                self.__old_excel_file_exists = True
        except FileNotFoundError:
            log.error(
                f'FileNotFoundError error occurred while reading {self.__name} - old excel file {self.__old_excel_file_sheet_name}')
        except ValueError:
            log.error(
                f'ValueError error occurred while reading {self.__name} - old excel file sheet {self.__old_excel_file_sheet_name}')
        except AttributeError:
            log.error(
                f'AttributeError error occurred while reading {self.__name} - old excel file sheet {self.__old_excel_file_sheet_name}')
        except ImportError:
            log.error(f'Please install missing Python Libraries.')

    def read_new_excel_file(self, log):
        log.info(
            f'Reading {self.__name} - new data file sheet {self.__new_excel_file_sheet_name} at {self.__new_excel_file_path}')
        try:
            if self.__old_excel_file_extension == Vendor.XLS_FILE_EXT or self.__old_excel_file_extension == Vendor.XLSX_FILE_EXT:
                self.__new_excel_file_data_frame = pd.read_excel(self.__new_excel_file_path,
                                                                 self.__new_excel_file_sheet_name)
                self.__new_excel_file_exists = True
                # print(self.__new_excel_file_data_frame)
            elif self.__old_excel_file_extension == Vendor.CSV_FILE_EXT:
                self.__new_excel_file_data_frame = pd.read_csv(self.__new_excel_file_path)
                self.__new_excel_file_exists = True
        except FileNotFoundError:
            log.error(
                f'FileNotFoundError error occurred while reading {self.__name} - new excel file {self.__new_excel_file_sheet_name}')
        except ValueError:
            log.error(
                f'ValueError error occurred while reading {self.__name} - new excel file sheet {self.__new_excel_file_sheet_name}')
        except AttributeError:
            log.error(
                f'AttributeError error occurred while reading {self.__name} - new excel file sheet {self.__new_excel_file_sheet_name}')
        except ImportError:
            log.error(f'Please install missing Python Libraries.')

    def compare_data_frames(self, log):
        if (self.__old_excel_file_exists and self.__new_excel_file_exists):
            columns = self.__new_excel_file_data_frame.columns
            self.__old_excel_file_data_frame['version'] = "Old"
            self.__new_excel_file_data_frame['version'] = "New"
            combined_data_frame = pd.concat([self.__old_excel_file_data_frame, self.__new_excel_file_data_frame],
                                            ignore_index=True)
            final_data_frame = self.remove_duplicate_rows(combined_data_frame, columns)

            if self.__external_id_postfix is not None:
                log.info('Postfix column is not empty, therefore adding an extra column for an external ID.')
                final_data_frame['External ID'] = str(
                    final_data_frame[self.__look_up]) + '-' + self.__external_id_postfix
            else:
                log.info('Postfix column is empty, therefore not adding an extra column for an external ID.')
            log.info(f'Saving CSV file for {self.__name} having {len(final_data_frame)} rows at {self.__results_csv_file_path}')

            final_data_frame.to_csv(self.__results_csv_file_path, index=False)
            self.upload_result_excel_file_to_ftp_server(log)

            try:
                log.info(f'Removing the old file from {self.__old_excel_file_path}.')
                os.remove(self.__old_excel_file_path)
                log.info(f'Removed the old file from {self.__old_excel_file_path}.')
            except OSError as e:
                log.info(f'Unable to remove old file from {self.__old_excel_file_path}, {e.strerror} occurred.')

            try:
                log.info(f'Renaming the new file at {self.__new_excel_file_path}.')
                os.rename(self.__new_excel_file_path, self.__old_excel_file_path)
                log.info(f'Renamed the new file at {self.__new_excel_file_path}.')
            except OSError as e:
                log.info(f'Unable to rename new file at {self.__new_excel_file_path}, {e.strerror} occurred.')

        else:
            log.error(f'Unable to compare the {self.__name} old and new excel files.')

    def remove_duplicate_rows(self, _combined_data_frame, _columns):
        changes = _combined_data_frame.drop_duplicates(subset=_columns, keep='last')
        duplicate_rows = changes[changes[self.__look_up].duplicated() == True][self.__look_up].tolist()
        duplicates = changes[changes[self.__look_up].isin(duplicate_rows)]
        cols = self.__changes
        cols.append(self.__look_up)
        duplicates = duplicates.drop_duplicates(cols, keep=False)
        duplicates = duplicates[(duplicates["version"] == "New")]
        duplicates = duplicates.drop(['version'], axis=1)
        return duplicates

    def download_source_excel_file_from_ftp_server(self, log):
        filename = f'{self.__source_ftp_filename}{self.__new_excel_file_extension}'
        try:
            ftp = FTP(self.__source_ftp_url)
            ftp.login(user=self.__source_ftp_user, passwd=self.__source_ftp_pass)
            ftp.cwd(self.__source_ftp_path)
            with open(self.__new_excel_file_path, "wb") as file:
                # use FTP's RETR command to download the file
                ftp.retrbinary(f'RETR {filename}', file.write)
            # file = open(self.__new_excel_file_path, 'wb')  # file to send
            # ftp.retrbinary('RETR ' + filename, file.write)  # send the file
            self.__new_ftp_file_exists = True
            ftp.quit()
        except all_errors as err:
            self.__new_ftp_file_exists = False
            log.error(f'Unable to download the new file {filename} from ftp {self.__source_ftp_url}.')

    def upload_result_excel_file_to_ftp_server(self, log):
        try:
            ftp = FTP(self.__results_ftp_url)
            ftp.login(user=self.__results_ftp_user, passwd=self.__results_ftp_pass)
            print(self.__results_ftp_path)
            ftp.cwd(self.__results_ftp_path)
            with open(self.__results_csv_file_path, 'rb') as file:
                # use FTP's STOR command to upload the file
                ftp.storbinary(f'STOR {os.path.basename(file.name).strip()}', file)
            # file = open(self.__results_csv_file_path, 'rb')  # file to send
            # ftp.storbinary('STOR ' + os.path.basename(file.name).strip(), file)  # send the file
            ftp.quit()
        except all_errors as err:
            log.error(f'Unable to upload the results file {self.__results_csv_file_path} to ftp {self.__results_ftp_url}.')
