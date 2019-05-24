from bs4 import BeautifulSoup
import requests
from fake_useragent import UserAgent
import re
from docx2csv import extract_tables, extract
from glob import glob
import re
import os
import win32com.client as win32
from win32com.client import constants
import pandas as pd

source_path = 'C:\\Users\\MikhaylovAV1\\Documents\\GKS_forms_1\\source\\'
paths = glob('C:\\Users\\MikhaylovAV1\\Documents\\GKS_forms_1\\source\\*.doc', recursive=True)
path_docx = glob('C:\\Users\\MikhaylovAV1\\Documents\\GKS_forms_1\\docx_formated\\*.docx', recursive=True)
path_xlsx = glob('C:\\Users\\MikhaylovAV1\\Documents\\GKS_forms_1\\excel\\*.xlsx', recursive=True)


def load_data(page):
    url = 'http://www.gks.ru/metod/form19/Page%d.html' % (page)
    r = requests.get(url, headers={'User-Agent': UserAgent().random})
    r.encoding = 'windows-1251'
    return r.text


def parser(text):
    soup = BeautifulSoup(text,'lxml')
    td = soup.find_all('td')
    for i in td:
        try:
            if i.nextSibling.get_text().split(' ')[0] in ('Сведения', 'Обследование', 'Показатели', 'Основные'):
                print('http://www.gks.ru/metod/form19' + i.a.get('href')[1:])
                r = requests.get('http://www.gks.ru/metod/form19' + i.a.get('href')[1:])
                name = i.a.get_text() + '.doc'
                if not os.path.exists(source_path):
                    os.makedirs(source_path)
                with open(source_path + name, 'wb') as f:
                    f.write(r.content)
            else:
                pass
        except AttributeError:
            pass



def save_as_docx(path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate ()
    new_file_abs = os.path.abspath(path[:len(path)-18] + '\\docx_formated\\' + path[len(path)-11:])
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)



def concat_xlsx(paths):
    excels = [pd.ExcelFile(name) for name in paths]
    frames = [x.parse(x.sheet_names[0], header=None, index_col=None, names=['Pokazatel', 'Attribute', 'SMTH', 'form'],converters={'form':str}) for x in excels]
    #frames[1:] = [df[1:] for df in frames[1:]]
    frames = [df for df in frames]
    combined = pd.concat(frames)
    combined.columns = ['Pokazatel', 'Attribute', 'SMTH', 'form']
    #combined.dropna(inplace=True)
    combined.to_excel("C:\\Users\\MikhaylovAV1\\Documents\\GKS_forms_1\\Result\\Result.xlsx", header=True, index=False)



if __name__ == '__main__':
    while True:
        print('1 - Собрать формы с сайта Росстата')
        print('2 - Перевести формы в формат docx')
        print('3 - Вытащить таблицы и поместить их в Excel')
        print('4 - Собрать все в один файл')
        print('0 - Выход')
        print()
        var = int(input('Введите число, соответствующее команде: '))
        try:
            if var == 1:
                for i in range(2, 34):
                    if i in (12, 16, 17, 24):
                        continue
                    else:
                        parser(load_data(i))
                print('Формы успешно собраны')
                print()
                continue
            elif var == 2:
                print('Этот процесс может длиться около 2 минут, ожидайте')
                for path in paths:
                    save_as_docx(path)
                print('Формы успешно сохранены в формате docx')
                print()
                continue
            elif var == 3:
                for path in path_docx:
                    try:
                        try:
                            extract(filename=path, format="xlsx", sizefilter=6,
                                    singlefile=True)
                        except UnboundLocalError:
                            try:
                                extract(filename=path, format="xlsx", sizefilter=5,
                                        singlefile=True)
                            except UnboundLocalError:
                                extract(filename=path, format="xlsx", sizefilter=4,
                                        singlefile=True)
                    except:
                        pass
                print('Таблицы помещены в xlsx файлы')
                print()
                continue
            elif var == 4:
                concat_xlsx(path_xlsx)
                print('Данные собраны в один файл')
                continue
            else: break
        except ValueError:
            print('Неверное знанчение, выберите 1, 2, или 3')
        except PermissionError:
            print('Закройте открыте файлы')




    # for i in range (1,34):
    #     if i in (12, 16, 17, 24):
    #         continue
    #     else:
    #         parser(load_data(i))
    # for path in paths:
    #     save_as_docx(path)
    # for path in path_docx:
    #     try:
    #         try:
    #             extract(filename=path, format="xlsx", sizefilter=6,
    #                     singlefile=True)
    #         except UnboundLocalError:
    #             try:
    #                 extract(filename=path, format="xlsx", sizefilter=5,
    #                         singlefile=True)
    #             except UnboundLocalError:
    #                 extract(filename=path, format="xlsx", sizefilter=4,
    #                         singlefile=True)
    #     except:
    #         pass
    # concat_xlsx(path_xlsx)
