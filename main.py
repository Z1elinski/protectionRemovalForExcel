import os
import zipfile
import PySimpleGUI as PySG
# import shutil


class ProtectionRemoval:

    def __init__(self, path, new_name=None):
        """
        :param path: path of a xlsx. file to be unlocked
        :param new_name: name of unlocked file
        """
        self.path = path.replace('/', '\\')
        # new name can be inserted but here is default
        if new_name is None:
            self.new_name = 'unlocked_' + os.path.basename(path).removesuffix('.xlsx')
        else:
            self.new_name = new_name

    @staticmethod
    def file_first_checker(path):
        """
        Checks if path leads to .xlsx file
        :return: bool
        """
        if os.path.exists(path) and path.endswith('.xlsx'):
            return True
        else:
            return False

    def change_xlsx_to_zip(self):
        """
        Changes file extension from .xlsx to .zip
        """
        try:
            os.rename(self.path, self.path.replace('.xlsx', '.zip'))
            print('Extension changed from .xlsx to .zip')
            self.path = self.path.replace('.xlsx', '.zip')
            print('new path: {}'.format(self.path))
        except WindowsError as e:
            print('Wrong path or extension: {}'.format(e))

    def change_zip_to_xlsx(self):
        """
        Changes file extension from .zip to .xlsx
        """
        try:
            os.rename(self.path, self.path.replace('.zip', '.xlsx'))
            print('Extension changed from .zip to .xlsx')
            self.path = self.path.replace('.zip', '.xlsx')
            print('new path: {}'.format(self.path))
        except WindowsError as e:
            print('Wrong path or extension: {}'.format(e))

    def unpack_zip(self):
        """
        Unpacks zip file to folder in path directory
        """
        try:
            with zipfile.ZipFile(self.path, 'r') as archive:
                archive.extractall(self.path.replace('.zip', ''))
                print('{} extracted'.format(os.path.basename(self.path)))
                self.path = self.path.replace('.zip', '')
                print('current directory: {}'.format(self.path))
        except FileNotFoundError as e:
            print('Wrong path or extension: {}'.format(e))

    def pack_zip(self):
        """
        Packs files in path directory to an archive
        """
        try:
            # self.path[:self.path.rfind('\\')] drops directory from which zip will be build.
            with zipfile.ZipFile(self.path[:self.path.rfind('\\')] + '\\' + self.new_name + '.zip', 'w') as archive:
                for directory, folders, files in os.walk(self.path):
                    for folder in folders:
                        abs_name = os.path.abspath(os.path.join(directory, folder))
                        arc_name = abs_name[len(self.path) + 1:]
                        archive.write(os.path.join(directory, folder), arcname=arc_name)
                    for file in files:
                        abs_name = os.path.abspath(os.path.join(directory, file))
                        arc_name = abs_name[len(self.path) + 1:]
                        archive.write(os.path.join(directory, file), arcname=arc_name)
            print('Created .zip file of: {} content'.format(self.path))
            # Changes path to new file, so it can be changed back to .xlsx
            self.path = self.path[:self.path.rfind('\\')] + '\\' + self.new_name + '.zip'
        except WindowsError as e:
            print('Error: {}'.format(e))
            # os.remove(self.path[:self.path.rfind('\\')] + '\\' + self.new_name + '.zip')

    def get_str_from_xml(self):
        """
        Opens xml file and replaces it with unprotected one
        :return:
        """
        xml_path = self.path + r'\xl\worksheets\sheet1.xml'
        try:
            with open(xml_path, 'r') as file:
                xml_str = file.read()
            new_xml_str = ProtectionRemoval.remove_protection_string(xml_str)
            with open(xml_path, 'w') as file:
                file.write(new_xml_str)
            print('Content of sheet1.xml replaced')
        except WindowsError as e:
            print(e)

    @staticmethod
    def remove_protection_string(str_text):
        """
        Removes part of string responsible for file protection
        :param str_text: string from .xml file
        :return: string without part of the code
        """
        protection_index_start = str_text.find('sheetProtection') - 1
        protection_index_end = str_text[protection_index_start:].find('>') + protection_index_start + 1
        # print(protection_index_start, protection_index_end)
        # print(str_text[protection_index_start:])
        # print(str_text[protection_index_start:protection_index_end])
        str_to_remove = str_text[protection_index_start:protection_index_end]
        return str_text.replace(str_to_remove, '')


if __name__ == '__main__':
    # fileToChangePath = r'C:\Users\zieli\Desktop\szybki test\50149047 - Tabela techniczna_poz.2.xlsx'
    # fileToRemoveProtection = ProtectionRemoval(fileToChangePath)
    # fileToRemoveProtection.change_xlsx_to_zip()
    # fileToRemoveProtection.unpack_zip()
    # fileToRemoveProtection.get_str_from_xml()
    # fileToRemoveProtection.pack_zip()
    # fileToRemoveProtection.change_zip_to_xlsx()

    layout = [
        [PySG.Text('Excel file path'), PySG.In(size=(40, 1), enable_events=True, key='-FILE-'), PySG.FileBrowse(),
         PySG.Button('OK')],
        [PySG.Text('Output'), PySG.Output(size=(80, 5))]
        ]
    window = PySG.Window('Protection removal for excel files', layout)

    fileToChangePath = ''

    while True:
        event, value = window.read()

        if event == 'Exit' or event == PySG.WINDOW_CLOSED:
            break

        if event == '-FILE-':
            fileToChangePath = value['-FILE-']

        if event == 'OK':
            if ProtectionRemoval.file_first_checker(fileToChangePath):
                print('Path OK')
                window.refresh()
                fileToRemoveProtection = ProtectionRemoval(fileToChangePath)
                fileToRemoveProtection.change_xlsx_to_zip()
                fileToRemoveProtection.unpack_zip()
                fileToRemoveProtection.get_str_from_xml()
                fileToRemoveProtection.pack_zip()
                fileToRemoveProtection.change_zip_to_xlsx()
            else:
                print('Wrong path')
            window.refresh()

    window.close()
