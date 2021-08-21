import os
import pathlib
import zipfile
import xml.etree.ElementTree
import pandas as pd
import requests as requests
from docx import Document

"""Константы"""
USERS = ('Бутакова', 'Винтова', 'Грибанова', 'Чернова')
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
CELL = WORD_NAMESPACE + 'tc'


class ParserFile:
    @staticmethod
    def open_file(file):
        try:
            with zipfile.ZipFile(file) as docx:
                return xml.etree.ElementTree.XML(docx.read('word/document.xml'))
        except zipfile.BadZipFile:
            return False

    @staticmethod
    def get_all_text(file):
        list_text = []
        for text in file.iter(PARA):
            item = ''.join(node.text for node in text.iter(TEXT))
            if item == '' or item == ' ' or item.__len__() < 2:
                continue
            else:
                list_text.append(item)
        return list_text

    @staticmethod
    def find_tasks(text):
        support_param = True
        for el in text:
            if support_param:
                if el[:2] == '2.':
                    index_start_task = text.index(el)
                    support_param = False
            if el[:25] == 'Настоящий Отчет по резуль':
                index_stop_task = text.index(el)
                break
        return text[index_start_task:index_stop_task]

    @staticmethod
    def get_task(tasks):
        alphabet = 'Й Ц У К Е Н Г Ш Щ З Х Ъ Ф Ы В А П Р О Л Д Ж Э Ё Я Ч С М И Т Ь Б Ю' \
                   'й ц у к е н г ш щ з х ъ ф ы в а п р о л д ж э ё я ч с м и т ь б ю'
        list_tasks = []
        for task in tasks:
            i = 0
            while i < task.__len__():
                symbol = task[i:i + 1]
                if symbol in alphabet:
                    list_tasks.append(task[i:])
                    break
                else:
                    i += 1
        task_return = []
        for reverse in list_tasks:
            task_return.append(reverse[::-1])

        return task_return

    @staticmethod
    def get_text_in_task(text, task, task_end):
        pass

    @staticmethod
    def get_tables(file):
        tables = []
        tables_all_text = []
        i = 0
        for table in file.iter(TABLE):
            table_text = []
            text_a = []
            for row in table.iter(ROW):
                text = []
                for cell in row.iter(CELL):
                    text.append(''.join(node.text for node in cell.iter(TEXT)))
                    text_a.append(''.join(node.text for node in cell.iter(TEXT)))
                if text.__len__() > 1:
                    table_text.append(text)
            tables_all_text.append(text_a)
            if table_text.__len__() > 0:
                i += 1
                tables.append(table_text)

        return tables, tables_all_text


def work_from_text(tasks, text):
    i = 1
    task_and_text = []
    for task, task_next in zip(tasks, tasks[i:]):
        try:
            index_start = text.index(task)
        except ValueError:
            try:
                index_start = index_stop
            except:
                for el in text:
                    if task_next in el:
                        index_start = text.index(el)
                        break
        try:
            index_stop = text.index(task_next)
        except ValueError:
            for el in text[index_start:]:
                if task_next in el:
                    index_stop = text.index(el)
                    break
        text_task = text[index_start:index_stop]
        i += 1
        task_and_text.append({'task': task, 'text': text_task})
    a = tasks[-1]
    try:
        task_and_text.append({'task': tasks[-1], 'text': text[text.index(a):]})
    except ValueError:
        pass

    return task_and_text


def send_request(user, task, project):
    for el in task:
        TEXT = ''
        for text in el['text']:
            TEXT += f'\n\t{text}'

        response = requests.post('http://127.0.0.1:8000/api-create-task/',
                                 data={
                                     "user": user,
                                     "task": el['task'],
                                     "text": TEXT,
                                     'project': project.replace('Отчет', "")
                                 })
        if response.status_code == 200:
            print(f'Отправлена задача')


def send_file_request(name, task, text, project, user):
    file_ob = {'uploaded_file': open(name, 'rb')}
    response = requests.post('http://127.0.0.1:8000/api-past-file/',
                             files=file_ob,
                             data={
                                 "user": user,
                                 "task": task,
                                 "text": text,
                                 'project': project.replace('Отчет', "")
                             })
    if response.status_code == 200:
        print(f'Отправлен файл в задачу {task[:30]}')
        os.remove(name)


def create_table_and_request(table1, task, text, project, user):
    data = pd.DataFrame(table1)
    document = Document()

    document.add_heading(task)
    table = document.add_table(rows=(data.shape[0]), cols=data.shape[1])  # First row are table headers!
    for i, column in enumerate(data):
        for row in range(data.shape[0]):
            table.cell(row, i).text = str(data[column][row])
    table.style = 'TableGrid'
    name = task[:30].replace('Отчет', "") + '.docx'
    document.save(name)

    send_file_request(name, task, text, project, user)


def work_from_text_and_tables(tuple_task, tables, project, user):
    for text in tuple_task:
        for table, table_app in zip(tables[1], tables[0]):
            b = table
            a = text['text']
            result = list(set(a) & set(b))
            if result.__len__() == table.__len__():
                c = list(set(text['text'])-set(result))
                create_table_and_request(table_app, text['task'], c, project, user)


def work_in_file(dir, user):
    for currentFile in dir.iterdir():
        treeFile = ParserFile.open_file(currentFile)  # Получаем дерево обьектов файла
        if not treeFile:
            continue
        all_text = ParserFile.get_all_text(treeFile)  # Получаем весь текст файла
        all_tasks_notclear = ParserFile.find_tasks(all_text)  # Поиск всех задач из файла

        all_task_reverse = ParserFile.get_task(all_tasks_notclear)  # Удаление нумерации задачи и ее разворот
        all_task = ParserFile.get_task(all_task_reverse)  # Удаление нумерации страниц и разворот братно

        TaskText = work_from_text(all_task, all_text)  # Соединение задачи и текста
        tables = ParserFile.get_tables(treeFile)  # Получение всех таблиц
        # send_request(user, TaskText, currentFile.name)
        work_from_text_and_tables(TaskText, tables, currentFile.name, user)  # Соединение таблицы и задачи

        print(1)


def start():
    for user in USERS:
        currentDirectory = pathlib.Path(f'./files/{user}/')  # Открываем директорию с файлами
        work_in_file(currentDirectory, user)  # Запускаем цикл по каждому файлу в директории


if __name__ == '__main__':
    start()
