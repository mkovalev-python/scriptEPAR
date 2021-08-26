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
    def clear_num_page_in_task(task):
        task = task[::-1]
        index = 0
        for el in task:
            try:
                int(el)
                index += 1
            except ValueError:
                break
        return task[index:][::-1]

    @staticmethod
    def tree_tasks(tasks):
        task_tree_list = []
        for task in tasks:
            task_split = task.split('.')
            index = 0
            header1 = []

            for el in task_split:
                try:
                    int(el)
                    index += 1
                except ValueError:
                    break
            if index == 1:
                header2 = []
                header3 = []
                header4 = []
                header1.append(task[2:])
                task_tree = {'H1': header1}
                task_tree_list.append(task_tree)
            if index == 2:
                del task_tree_list[-1]
                header2.append(task[4:])
                task_tree.update({'H2': header2})
                task_tree_list.append(task_tree)
            if index == 3:
                del task_tree_list[-1]
                header3.append(task[6:])
                task_tree.update({'H3': header3})
                task_tree_list.append(task_tree)
            if index == 4:
                del task_tree_list[-1]
                header4.append(task[8:])
                task_tree.update({'H4': header4})
                task_tree_list.append(task_tree)

        return task_tree_list


def work_in_file(dir, user):
    for currentFile in dir.iterdir():
        treeFile = ParserFile.open_file(currentFile)  # Получаем дерево обьектов файла
        if not treeFile:
            continue
        all_text = ParserFile.get_all_text(treeFile)  # Получаем весь текст файла
        all_tasks_notclear = ParserFile.find_tasks(all_text)  # Поиск всех задач из файла
        tasks = []
        for task in all_tasks_notclear:
            tasks.append(ParserFile.clear_num_page_in_task(task))
        task_tree = ParserFile.tree_tasks(tasks)  # Создание дерева задач
        print(1)


def start():
    for user in USERS:
        currentDirectory = pathlib.Path(f'./files/{user}/')  # Открываем директорию с файлами
        work_in_file(currentDirectory, user)  # Запускаем цикл по каждому файлу в директории


if __name__ == '__main__':
    start()
