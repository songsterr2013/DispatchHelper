from os import remove, listdir
import os.path
import sys


class DeleteFiles:

    def __init__(self):

        self.root_path = ''

        if getattr(sys, 'frozen', False):  # 如果為exe
            self.root_path = os.path.dirname(sys.executable)
        elif __file__:
            self.root_path = os.path.dirname(__file__)

        self.previous_root_path = os.path.dirname(self.root_path)  # 取得檔案所在的前一路徑

    def del_1(self):

        load_open_path_1 = os.path.join(self.previous_root_path, "1_ERP產品_excel")

        try:
            load_file_1 = os.path.join(load_open_path_1, listdir(load_open_path_1)[0])

            try:
                if load_file_1[-4:] == 'xlsx' or self.load_file_1[-3:] == 'xls':
                    remove(load_file_1)
            except OSError as e:
                print(e)

        except IndexError as e:
            print('資料夾中沒有檔案:', e)

    def del_2(self):

        load_open_path_2 = os.path.join(self.previous_root_path, "2_炸製令結果_excel")

        try:
            load_file_2 = os.path.join(load_open_path_2, listdir(load_open_path_2)[0])

            try:
                if load_file_2[-4:] == 'xlsx' or self.load_file_2[-3:] == 'xls':
                    remove(load_file_2)
            except OSError as e:
                print(e)

        except IndexError as e:
            print('資料夾中沒有檔案:', e)

    def del_3(self):

        load_open_path_3 = os.path.join(self.previous_root_path, "3_派工單_excel")

        try:
            load_file_3 = os.path.join(load_open_path_3, listdir(load_open_path_3)[0])

            try:
                if load_file_3[-4:] == 'xlsx' or self.load_file_3[-3:] == 'xls':
                    remove(load_file_3)
            except OSError as e:
                print(e)

        except IndexError as e:
            print('資料夾中沒有檔案:', e)


if __name__ == '__main__':
    delete_files = DeleteFiles()
    delete_files.del_1()
    delete_files.del_2()
    delete_files.del_3()
