# coding: utf8
import os
import xlrd


class XLS(object):
    """
    Анализ xls файлов
    """
    def __init__(self, dir='xlsFiles'):
        self.__dir = dir

    def __isDir(self):
        """
        Проверка на наличе деректории
        """
        try:
            self.__dirFile = os.listdir(self.__dir)
        except:
            raise  TypeError('{} папка не найдена'.format(self.__dir))

    def analysis(self):
        """
        Обход по файлам в деректориии
        :return dict:
        """
        self.__isDir()
        data = {}
        for file in self.__dirFile:
            form = os.path.splitext(file)[1]
            if not form == '.xls':
                print 'Фаил {} не подходит. Формат файла должен иметь расширение xls'.format(file)
            else:
                result = dict(self.file(file))
                data = result if len(data) == 0 else self.dictPlus(data, result)
        return data

    def file(self, file):
        """
        Обход по файлу
        :param file Фаил для анализа:
        :return list:
        """
        try:
            file = '%s/%s'%(self.__dir, file)
            rb = xlrd.open_workbook(file, formatting_info=False, on_demand=True) # открытие файла
            sheet = rb.sheet_by_index(0) # считывание информации из файла
            worker, result = [], []
            value = sheet.row_values(0)[4:]
            for i in range(len(value))[::2]:
                worker.append(self.longest_common_substring(value[i], value[i+1]).strip()) # Получение имен сотрудниеков
            for i in range(1, sheet.nrows):
                info = self.fileRow(sheet.row_values(i)[4:])
                # формирование списка результов анализа по строкам
                result = info if len(result) == 0 else list(map(lambda a, b: a + b, result, info))
            return zip(worker, result)
        except:
            raise TypeError('Не удалось получить информацию из файла {}'.format(file))

    def fileRow(self, data):
        """
        Анализ строк файла
        :param data Строка файла:
        :return list:
        """
        info = []
        for i in range(len(data))[::2]:
            data[i] = [0, data[i]][self.is_number(data[i])]
            data[i+1] = [0, data[i+1]][self.is_number(data[i+1])]
            info.append(data[i+1] - data[i]) # Вычитаем из фактического количества дней планируемое
        return info

    def dictPlus(self, d1, d2):
        """
        Сложение двух словарей с сохранением уникальных элементов и сложением повторяющихся
        :param d1:
        :param d2:
        :return dict:
        """
        l1= list(map(lambda a: [a,d2[a]+d1[a]] if a in d2 else [a,d1[a]], d1.keys()))
        l2 = list(map(lambda a: [a,d2[a]+d1[a]] if a in d1 else [a,d2[a]], d2.keys()))
        return dict(l1+l2)


    def is_number(self, str):
        """
        Проверка на возможность преобразование строки в число
        :param str Строка для проверки:
        :return bool:
        """
        try:
            float(str)
            return True
        except ValueError:
            return False

    def longest_common_substring(self, s1, s2):
        """
        Получение общей подстроки у двух строк
        :param s1 Строка:
        :param s2 Строка:
        :return string:
        """
        m = [[0] * (1 + len(s2)) for i in range(1 + len(s1))]
        longest, x_longest = 0, 0
        for x in range(1, 1 + len(s1)):
            for y in range(1, 1 + len(s2)):
                if s1[x - 1] == s2[y - 1]:
                    m[x][y] = m[x - 1][y - 1] + 1
                    if m[x][y] > longest:
                        longest = m[x][y]
                        x_longest = x
                else:
                    m[x][y] = 0
        return s1[x_longest - longest: x_longest]