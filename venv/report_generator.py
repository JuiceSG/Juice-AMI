import io
import os
import docx
import logging
import itertools
import shotgun_api3
import requests
import numpy as np
import datetime
from dotenv import load_dotenv
from docx.oxml.shared import OxmlElement, qn
from docx.shared import Mm, Pt


load_dotenv('C:/Juice Pipeline/config.env')
sg_token = os.getenv('SG_AMI_FLASK_TOKEN')
sg_address = os.getenv('SG_ADDRESS')
sg_script_name = os.getenv('SG_AMI_FLASK_NAME')
logging.basicConfig(level=logging.DEBUG)


class GroupsValues:
    def __init__(self, post_dict, sg):
        self.__sg = sg
        self.__post_dict = post_dict
        self.__is_grouped = False
        self.__group_by = self.__get_grouping_columns()     # zawiera kody kolumn wedlug ktorych beda grupowane dane
        self.__groups = self.__get_groups()     # zawiera wszystkie mozliwe kombinacje grup
        self.__invisible_groups = self.__get_invisible_groups()

    @property
    def is_grouped(self):
        return self.__is_grouped

    @property
    def groups(self):
        return self.__groups

    @property
    def group_by(self):
        return self.__group_by

    @property
    def invisible(self):
        return self.__invisible_groups

    def __get_groups(self):
        if self.__group_by:
            groups = self.__get_groups_values(self.__group_by)
            groups = self.__all_combinations(groups)
            self.__is_grouped = True
            return groups
        else:
            return None

    def __get_groups_values(self, grouping_columns):
        groups_values = []
        for group_column in grouping_columns:
            group_values = self.__get_group_values(group_column)
            groups_values.append(group_values)
        return groups_values

    def __get_group_values(self, grouping_column):
        project_id = int(self.__post_dict['project_id'])
        project = self.__sg.find_one('Project', [['id', 'is', project_id]])
        entity_type = self.__post_dict['entity_type']
        entity_to_group_by = self.__sg.find(entity_type, [['project', 'is', project]], [grouping_column])
        groups = []
        for group in entity_to_group_by:
            if type(group[grouping_column]) is dict:
                if 'name' in group[grouping_column].keys():
                    group = group[grouping_column]['name']
                elif 'code' in group[grouping_column].keys():
                    group = group[grouping_column]['code']
            else:
                group = group[grouping_column]
            if group not in groups:
                groups.append(group)
        return groups

    def __get_grouping_columns(self):
        if 'grouping_columns' in self.__post_dict:
            grouping_columns = self.__post_dict['grouping_columns']
        elif 'grouping_column' in self.__post_dict:
            grouping_columns = self.__post_dict['grouping_column']
        else:
            return None
        grouping_columns = self.__convert_to_list(grouping_columns)
        return grouping_columns

    def __get_invisible_groups(self):
        columns = self.__post_dict['cols']
        columns = self.__convert_to_list(columns)
        invisible_groups = []
        if self.__group_by:
            for column in self.group_by:
                if column not in columns:
                    invisible_groups.append(column)
        return invisible_groups

    @staticmethod
    def __convert_to_list(data):
        converted_list = data.split(',')
        return converted_list

    @staticmethod
    def __all_combinations(groups):
        groups = itertools.product(*groups)
        groups = list(groups)
        return groups


class EntityData:
    """
    groups - grupy dla danego zapytania np okreslenie konretnego epizodu i podlegajacemu
    """
    def __init__(self, post_dict, groups, sg):
        self.__sg = sg
        self.__groups = groups
        self.__post_dict = post_dict
        self.__steps_dict = self.__get_entity_steps_dict()
        self.__steps_code_list = list(map(lambda x: x['code'], self.__steps_dict))
        self.__steps_short_code_list = list(map(lambda x: x['short_name'], self.__steps_dict))
        self.__data = self.__get_data_from_sg()
        self.__groups_columns_names = self.__get_groups_cols_names()
        self.__data_cols_names = self.__get_data_cols_names()

    @property
    def data(self):
        return self.__data

    @property
    def groups_columns(self):
        return self.__groups_columns_names

    @property
    def data_columns(self):
        return self.__data_cols_names

    def __get_entity_steps_dict(self):
        entity_type = self.__post_dict['entity_type']
        steps = self.__sg.find('Step', [['entity_type', 'is', entity_type]], ['code', 'short_name'])
        steps = map(lambda x: {'code': x['code'], 'short_name': x['short_name']}, steps)  # wyluskuje niezbedne dane
        steps = list(steps)
        return steps

    def __get_data_from_sg(self):
        project = self.__get_project()
        entity_type = self.__post_dict['entity_type']
        invisible_columns = self.__groups.invisible
        selected_columns = self.__convert_to_list('cols')
        selected_columns += invisible_columns
        order = self.__set_order()
        selected_ids = self.__get_selected_ids()
        filters = [
            ['project', 'is', project],
            ['id', 'in', selected_ids],
        ]
        data = self.__sg.find(entity_type, filters, selected_columns, order)
        data = self.__format_data(data)
        return data

    def __set_order(self):
        order_by = self.__convert_to_list('sort_column')
        order_direction = self.__convert_to_list('sort_direction')
        order = []
        for by, direction in zip(order_by, order_direction):
            order_dict = {'column': by, 'direction': direction}
            order.append(order_dict)
        return order

    def __get_project(self):
        project_id = int(self.__post_dict['project_id'])
        project = self.__sg.find_one('Project', [['id', 'is', project_id]])
        return project

    def __get_selected_ids(self):
        selected_ids = self.__post_dict['selected_ids']
        selected_ids = selected_ids.split(',')
        selected_ids = map(lambda x: int(x), selected_ids)
        return list(selected_ids)

    def __format_data(self, data):
        formatted_data = []
        for entity_dict in data:
            formatted_element = self.__format_entity(entity_dict)
            formatted_data.append(formatted_element)
        return formatted_data

    def __format_entity(self, entity_dict):
        entity = {}
        columns_display_names = self.__convert_to_list('column_display_names')      # tutaj dodac niewidoczne pola
        columns_display_names += self.__groups.invisible
        columns_codes = self.__convert_to_list('cols')      # tutaj dodac niewidoczne pola
        columns_codes += self.__groups.invisible
        for column_display_name, column_code in zip(columns_display_names, columns_codes):
            value = entity_dict[column_code]
            if type(value) is dict:
                value = self.__get_entity_name(value)
            if value is None:
                if column_display_name in self.__steps_code_list:  # wykonywanie dla ROZWINIETYCH podsumowan
                    value = self.__check_summary_column(column_display_name, entity_dict, is_expended=True)
                elif column_display_name in self.__steps_short_code_list:  # wykonywane dla ZWINIETYCH podsumowan)
                    value = self.__check_summary_column(column_display_name, entity_dict, is_expended=False)
            entity[column_display_name] = value
        return entity

    def __convert_short_code_to_name(self, column_display_name):
        step_name = None
        for step in self.__steps_dict:
            if column_display_name == step['short_name']:
                step_name = step['code']
        if not step_name:
            print('Problem ze stepami')
        return step_name

    def __convert_to_list(self, data):
        converted_list = self.__post_dict[data].split(',')
        return converted_list

    def __check_summary_column(self, column_code, entity_dict, **kwargs):
        is_expended = kwargs.get('is_expended')
        if not is_expended:
            column_code = self.__convert_short_code_to_name(column_code)
        entity = self.__get_entity(entity_dict)  # przechowuje dane aktualnie przetwarzanego elementu
        fields = ['step', 'content', 'sg_status_list', 'task_assignees', 'start_date', 'duration']
        entity_tasks = self.__sg.find('Task', [['entity', 'is', entity]], fields)
        entity_task_status = self.__check_tasks_status(entity_tasks, column_code)
        entity_task_status['is_expended'] = is_expended
        return entity_task_status

    def __get_groups_cols_names(self):
        if self.__groups.group_by is None:
            return None
        groups_display_name = []
        for group in self.__groups.group_by:
            group_display_name = self.__get_display_name(group)
            groups_display_name.append(group_display_name)
        return groups_display_name

    def __get_display_name(self, group):
        names_dict = self.__get_columns_names_dict()
        if group in names_dict.keys():
            group = names_dict[group]
        elif group in self.__groups.invisible:
            group = group
        return group

    def __get_columns_names_dict(self):
        """
        tworzy slownik nazw kolumn z code na display
        :return:
        """
        names_dict = {}
        displayed_columns_names = self.__convert_to_list('column_display_names')
        code_columns_names = self.__convert_to_list('cols')
        for code, display in zip(code_columns_names, displayed_columns_names):
            # print('%s = %s' % (code, display))
            names_dict[code] = display
        return names_dict

    def __get_data_cols_names(self):
        displayed_columns = self.__convert_to_list('column_display_names')
        if self.__groups.group_by is None:
            return displayed_columns
        groups = []
        for column in displayed_columns:
            if column in self.__groups_columns_names:
                pass
            else:
                groups.append(column)
        return groups

    def __get_entity(self, entity_dict):
        entity_type = entity_dict['type']
        entity_id = entity_dict['id']
        entity = self.__sg.find_one(entity_type, [['id', 'is', entity_id]])
        return entity

    @staticmethod
    def __check_tasks_status(entity_tasks, column_code):
        total_tasks = 0
        fin_tasks = 0
        step_tasks = []
        for task in entity_tasks:
            if task['step']['name'] == column_code:
                step_tasks.append(task)
                total_tasks += 1
                if task['sg_status_list'] == 'fin':
                    fin_tasks += 1
        if total_tasks > 0:
            step_status = float(fin_tasks) / float(total_tasks) * 100
            step_status = round(step_status)
        else:
            step_status = None
        step_summary = {
            'step_code': column_code,
            'tasks': step_tasks,
            'total_tasks': total_tasks,
            'fin_tasks': fin_tasks,
            'summary': step_status,
        }
        return step_summary

    @staticmethod
    def __get_entity_name(dictionary):
        if 'name' in dictionary:
            entity_name = dictionary['name']
        elif 'code' in dictionary:
            entity_name = dictionary['code']
        else:
            entity_name = dictionary
        return entity_name


class DocXFormatReport:
    # klasa odpowiedzialna za tworzneie dokumentu DOCX zwierajacaego zaznaczone elementy na stronie SG
    def __init__(self, data, groups, document_title):
        self.__document_title = document_title
        self.__data = data
        self.__groups = groups
        self.__doc = docx.Document()
        self.__doc_settings()
        self.__report = self.create_docx()

    @property
    def report(self):
        return self.__report

    def __doc_settings(self):
        section = self.__doc.sections[0]
        section.page_height = Mm(297)
        section.page_width = Mm(210)
        section.left_margin = Mm(5)
        section.right_margin = Mm(5)
        section.top_margin = Mm(5)
        section.bottom_margin = Mm(5)
        section.header_distance = Mm(5)
        section.footer_distance = Mm(5)
        style = self.__doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(8)

    def create_docx(self):
        self.__create_document_header()
        self.__create_data_table()
        # self.__doc.save('test.docx')
        file_stream = self.__file_stream()
        return file_stream

    def __create_document_header(self):
        header_table = self.__doc.add_table(rows=1, cols=2, style='TableNormal')
        header_table.autofit = False
        header_title = self.__document_title.split(' ')
        cell = header_table.cell(0, 0)
        col = header_table.columns[0].cells[0]
        col.width = Mm(40)
        paragraph = cell.paragraphs[0]
        run_image = paragraph.add_run()
        run_image.add_picture('c:/JuiceSG/AMI/img/J.png', height=Mm(20))
        paragraph = header_table.cell(0, 1).paragraphs[0]
        run_project_name = paragraph.add_run()
        run_project_name.add_break()
        run_project_name.add_text(header_title[0])
        run_project_name.font.size = Pt(14)
        run_project_name.bold = True
        run_project_name.add_break()
        run_entity_name = paragraph.add_run(header_title[1])
        run_entity_name.font.size = Pt(12)
        run_entity_name.bold = True
        self.__doc.add_paragraph('')  # zapewnia odstep miedzy tabelami

    def __create_data_table(self):
        if self.__data.groups_columns:
            columns = self.__data.groups_columns + self.__data.data_columns
        else:
            columns = self.__data.data_columns
        number_of_cols = len(columns)
        entity_table = self.__doc.add_table(rows=1, cols=number_of_cols, style='TableNormal')
        self.__create_entity_table_header(entity_table, columns)

        row_cells = entity_table.add_row().cells
        group_summary = self.__get_group_summary(self.__data.data)
        self.__add_group_summaries(group_summary, row_cells)

        self.__create_entities(entity_table)

    @staticmethod
    def __create_entity_table_header(table, columns):
        header_cells = table.rows[0].cells
        for cell, column_name in zip(header_cells, columns):
            paragraph = cell.paragraphs[0]
            run = paragraph.add_run()
            run.add_text(column_name)
            run.bold = True
        return True

    def __create_entities(self, table):
        if self.__data.groups_columns:
            groups = self.__groups.groups
            for group in groups:
                data_group = self.__group_data(group)
                self.__add_grouped_entities_rows(data_group, table)
        else:
            for data in self.__data.data:
                self.__add_entity_row(data, table)
        return True

    def __group_data(self, group):

        def group_filtering(data_entity):
            is_in_group = True
            for index, by in enumerate(self.__data.groups_columns):
                if data_entity[by] == group[index]:
                    is_in_group = True
                else:
                    is_in_group = False
                is_in_group *= is_in_group
                is_in_group = bool(is_in_group)
            return is_in_group

        grouped_data = list(filter(group_filtering, self.__data.data))
        if grouped_data:
            self.__remove_data(grouped_data)
        return grouped_data

    def __remove_data(self, grouped_data):
        for entity in grouped_data:
            self.__data.data.remove(entity)

    def __add_grouped_entities_rows(self, data_group, table):   # dodawanie danych dla grup
        self.__add_groups_headers(data_group, table)
        for data in data_group:
            self.__add_entity_row(data, table)
        return True

    def __add_entity_row(self, data, table):    # dodawanie wiersza danych dla okreslonego elementu z grupy
        row_cells = table.add_row().cells
        entity__index = self.__get_starting_entity_index()
        for index, key in enumerate(self.__data.data_columns):  #dodawanie danych do kolumn nie bedacych grupa
            cell = row_cells[entity__index + index]
            if (key == 'Thumbnail') and (data[key] is not None):
                self.__add_image_from_url(cell, data[key])
            elif type(data[key]) is dict:
                self.__add_data_from_step_column(data[key], cell)     # dla stepow przekazywane sa dane jako dict
            else:
                value = str(data[key])
                cell.text = value
        return True

    def __get_starting_entity_index(self):
        if self.__data.groups_columns:
            entity__index = len(self.__data.groups_columns)
        else:
            entity__index = 0
        return entity__index

    def __add_groups_headers(self, data_group, table):
        if data_group:
            for index, group in enumerate(self.__data.groups_columns):
                group_summary = self.__get_group_summary(data_group)
                row_cells = table.add_row().cells
                if group_summary:
                    self.__add_group_summaries(group_summary, row_cells)
                group_name = data_group[0]      # bierze pierwsze wystopienie z ktorego wyluskuje nazwy grup
                group_name = group_name[group]
                row_cells[index].text = str(group_name)
        return True

    def __add_group_summaries(self, group_summary, row_cells):
        entity__index = self.__get_starting_entity_index()
        for key, value in group_summary.items():
            if key in self.__data.data_columns:
                if value[1] == 0:
                    value = 0
                else:
                    value = float(value[1]) / float(value[0]) * 100.0
                    value = round(value)
                index = self.__data.data_columns.index(key)
                cell = row_cells[index + entity__index]
                paragraph = cell.paragraphs[0]
                run = paragraph.add_run()
                run.bold = True
                run.add_text('%s%%' % value)

    def __add_image_from_url(self, cell, image_url):
        margins = ['top', 'start', 'bottom', 'end']
        margin_size = '0'
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for margin in margins:
            node = OxmlElement('w:%s' % margin)
            node.set(qn('w:w'), margin_size)
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)
        img_width = Mm(40)
        image = self.__get_image_from_url(image_url)
        paragraph = cell.paragraphs[0]
        run = paragraph.add_run()
        run.add_picture(image, width=img_width)
        return True

    def __file_stream(self):
        file_stream = io.BytesIO()  # tworzy bufor w pamieci
        self.__doc.save(file_stream)    # zapisuje plik do buforu
        file_stream.seek(0)     # zeruje wskaznik pliku
        return file_stream

    @staticmethod
    def __get_group_summary(group_data):
        summary = {}
        for entity in group_data:
            for key, value in entity.items():
                if type(value) is dict:
                    if key in summary.keys():
                        total_tasks = summary[key][0]
                        fin_tasks = summary[key][1]
                    else:
                        total_tasks = 0
                        fin_tasks = 0
                    total_tasks += entity[key]['total_tasks']
                    fin_tasks += entity[key]['fin_tasks']
                    summary[key] = [total_tasks, fin_tasks]
        logging.debug('Podsumowanie ukonczenia taskow w grupie: %s' % summary)
        return summary

    @staticmethod
    def __get_image_from_url(url):
        response = requests.get(url, stream=True)
        image = io.BytesIO(response.content)  # obiekt wskazuje na miejsce w pamieci gdzie przechowywany jest obraz
        return image

    @staticmethod
    def __add_data_from_step_column(data, cell):
        if data['is_expended']:
            value = data['summary']
            """# warunek zrobiona jako tymczasowa trzeba bedzie ja dorobic na razie robi tabele wdlug standardu SG
            tasks = data['tasks']
            cell._element.clear_content()
            column_names = ('Status', 'Task name', 'Assigned to', 'Start Data', 'Duo data', 'Duration')
            table = cell.add_table(rows=1, cols=len(column_names))
            self.__create_entity_table_header(table, column_names)
            # dodawanie wierszy z danymi taskow
            for task in tasks:
                row_cells = table.add_row().cells
                row_cells[0] = task['sg_status_list']
                row_cells[1] = task['content']
                row_cells[2] = task['task_assignees']
                row_cells[3] = task['start_date']
                row_cells[4] = task['start_date']
                row_cells[5] = task['duration']
            # petla pozwalajaca na edycje rozmiaru czcionki dla calej tabeli
            for row in table.rows:
                for cell in row.cells:
                    paragraphs = cell.paragraphs
                    for paragraph in paragraphs:
                        for run in paragraph.runs:
                            font = run.font
                            font.size = Pt(6)"""
        else:
            value = data['summary']
        if value is None:
            cell.text = ''
        else:
            cell.text = '%s%%' % value
        return True


class ColumnsWidth:
    # klasa sluzaca do ustawienia szerkosci kolumn, bezcelowe dla Worda

    def __init__(self, entities):
        self.__page_width = 200
        self.__proportions = []
        self.get_columns_width(entities)

    @property
    def proportions(self):
        return self.__proportions

    @property
    def mm(self):
        mm = list(map(lambda x: round(x * self.__page_width), self.__proportions))
        return mm

    def get_longest_word(self, string):
        longest_word = string.split(' ')
        longest_word = max(longest_word)
        longest_word = len(longest_word)
        return longest_word

    def get_lengths(self, entities):
        cols_longest_words = []
        cols_max_length = []
        for entity in entities:
            longest_word = self.get_columns_longest_word(entity)
            max_length = self.get_columns_max_length(entity)
            cols_longest_words.append(longest_word)
            cols_max_length.append(max_length)
        return cols_longest_words, cols_max_length

    def get_columns_longest_word(self, entity):
        col_longest_word = []
        for column_name in entity:
            text = '%s %s' % (column_name, entity[column_name])
            longest_word = self.get_longest_word(text)
            if column_name in ['Thumbnail', 'thumbnail']:
                longest_word = longest_word / 10
            col_longest_word.append(longest_word)
        return col_longest_word

    def get_columns_width(self, entities):
        longest_words, max_lengths = self.get_lengths(entities)
        proportions = self.get_word_to_txt_proportions(longest_words, max_lengths)
        proportions = self.get_proportions(proportions)
        self.__proportions = proportions

    @staticmethod
    def get_columns_max_length(entity):
        cols_max_length = []
        for column_name in entity:
            text = '%s %s' % (column_name, entity[column_name])
            max_length = len(text)
            cols_max_length.append(max_length)
        return cols_max_length

    @staticmethod
    def get_proportions(length_proportions):
        avg_proportions = np.array(length_proportions)  # numpy uzyto w celu ogarniecia liczenia macierzy
        avg_proportions = avg_proportions.sum(axis=0)
        number_of_elements = len(avg_proportions)
        avg_proportions = list(map(lambda x: x / number_of_elements, avg_proportions))
        total = sum(avg_proportions)
        avg_proportions = list(map(lambda x: x / total, avg_proportions))
        return avg_proportions

    @staticmethod
    def get_word_to_txt_proportions(words_length, text_length):
        length_proportions = []
        for lenghts in zip(words_length, text_length):
            length_proportion = []
            index = 0
            while index < len(lenghts[0]):
                proportion = lenghts[1][index] / lenghts[0][index]
                index += 1
                length_proportion.append(proportion)
            length_proportions.append(length_proportion)
        return length_proportions


class ReportGenerator:
    def __init__(self, post_dict):
        self.__sg = shotgun_api3.Shotgun(sg_address, script_name=sg_script_name, api_key=sg_token)
        self.__post_dict = post_dict
        self.__project_name = post_dict['project_name']
        self.__entity_type = post_dict['entity_type']
        self.__document_title = '%s %ss.docx' % (self.__project_name, self.__entity_type.lower())

    @property
    def title(self):
        return self.__document_title

    def generate(self):
        groups = GroupsValues(self.__post_dict, self.__sg)
        logging.debug(' =========GroupsValues=========')
        logging.debug(' <groups> Nazwy grup: %s' % groups.groups)
        logging.debug(' <group_by> Kolumny wedlug ktorych maja byc grupowane dane: %s' % groups.group_by)
        logging.debug(' <invisible> Niewidoczne grupy: %s' % groups.invisible)
        data = EntityData(self.__post_dict, groups, self.__sg)
        logging.debug(' =========EntitiesData=========')
        logging.debug(' <data> Przekazywane dane: %s' % data.data)
        logging.debug(' <groups_columns> Wyswietlane nazwy kolumn z grupami: %s' % data.groups_columns)
        logging.debug(' <data_columns> Wyswietlane nazwy kolumn z danymi (bez grup): %s' % data.data_columns)
        logging.debug(' ==============================')
        new_report = DocXFormatReport(data, groups, self.__document_title)
        return new_report.report


def time_test(func):
    start = datetime.datetime.now()

    def wrapper(*args, **kwargs):
        func(*args, **kwargs)

    end = datetime.datetime.now()
    time = end - start
    func_name = func.__name__
    print('%s: %s' % (func_name, time))
    return wrapper
