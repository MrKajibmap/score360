import os
import shutil

import numpy as np
import pandas as pd
from pathlib import Path


def clear_folder(input_folder_path):
    for current_file_to_remove in os.listdir(Path(input_folder_path)):
        file_path = os.path.join(Path(input_folder_path), current_file_to_remove)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                print('Будет удален файл: ', file_path)
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                print('Будет удален файл: ', file_path)
                shutil.rmtree(file_path)
        except Exception as e:
            print('Ошибка при удалении - %s. Причина: %s' % (file_path, e))

def create_file(file_full_path):
    if not os.path.exists(file_full_path):
        print("Создание файла ", file_full_path)
        # Создаем файлы на основе пустого ДФ в 'Output_files/HardSkills/'
        df_new.to_excel(file_full_path)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    df_qna_general = []
    df_assessment_hard_general = []
    df_cross_general = []
    df_assessment_soft_general = []
    qna_transposed_general_pr = []
    if not Path('Input_files').is_dir():
        raise ValueError(f"[{Path('Input_files')}] не существует или не является директорией")

    print('ex_check: ', os.path.exists('Templates/Assessment_template.xlsx'))

    # Очистка директории с целевыми файлами по хард скиллам
    clear_folder('Output_files/HardSkills')
    clear_folder('Output_files/Cross')
    clear_folder('Output_files/SoftSkills')

    i = 0
    for file in Path('Input_files').iterdir():
        if file.name != '.DS_Store' and file.name != '':
            i = i + 1
            current_file_name = file.name.split('.')[0]
            # current_file_name = file.name
            print(f"Анализ выполняется для файла '{current_file_name}'.")
            # Достаем информацию по опроснику
            df_qna = pd.read_excel(file
                                   , nrows=5
                                   , skiprows=2
                                   , sheet_name='Оценка разработчиков'
                                   , header=None
                                   , names=['Question', 'Answer(Yes/No)', 'Comments']
                                   , usecols=range(2, 5))
            # df_qna['Employee'] = current_file_name
            # df_qna_general.append(df_qna)

            data = {
                'Желание быть куратором/наставником': str(df_qna.iloc[0, 1]),
                'Желание быть куратором/наставником--Comments': str(df_qna.iloc[0, 2]),
                "Желание проводить workshop/участие в IT Academy": str(df_qna.iloc[1, 1]),
                "Желание проводить workshop/участие в IT Academy--Comments": str(df_qna.iloc[1, 2]),
                "Желание углубить знания в предметных областях": str(df_qna.iloc[2, 1]),
                "Желание углубить знания в предметных областях -- Comments": str(df_qna.iloc[2, 1]),
                "Желание углубить знания в аналитике": str(df_qna.iloc[3, 1]),
                "Желание углубить знания в аналитике -- Comments": str(df_qna.iloc[3, 2]),
                "В свободной форме: какие скилы хотел бы прокачать": str(df_qna.iloc[4, 2])
            }
            cur_employee_qna = pd.DataFrame(data, index=[current_file_name])
            qna_transposed_general_pr.append(cur_employee_qna)
            print('Start hard for ', current_file_name)
            # достать информацию по оценке своей и коллег-разработчиков HARD
            df_assessment_hard = pd.read_excel(file
                                               , nrows=18
                                               , skiprows=15
                                               , sheet_name='Оценка разработчиков'
                                               , usecols=range(2, 23))
            list_of_columns = list(df_assessment_hard.columns)
            df_new = pd.DataFrame(index=range(0, 17))

            df_new.insert(0, "Критерий\Сотрудник", df_assessment_hard["Критерий\Сотрудник"])
            hard_iter = 0
            # находясь в файле, разбираем оценки как самого себя, так и коллег, пройдя по всем оценкам всех коллег (колонки = коллеги)
            df_for_list_employee = df_assessment_hard.loc[:, ~df_assessment_hard.columns.str.contains('^Unnamed')]
            df_for_list_employee = df_for_list_employee.loc[:, ~df_for_list_employee.columns.str.contains('^Критерий')]
            list_of_columns = list(df_for_list_employee.columns)
            # сюда вставить цикл по списку колонок (сотрудников)
            for current_column in list_of_columns:
                # print('Обработка колонки ', current_column, ' из ', df_assessment_hard.columns)
                hard_iter = hard_iter + 1
                curr_employee = str(list(df_assessment_hard.columns)[hard_iter])
                # Если разбираемая колонка = 'Критерий\Сотрудник', то пропускаем итерацию
                if str(curr_employee).upper() == str('Критерий\Сотрудник').upper():
                    # print("Пропускаем итерацию для колонки Критерий\Сотрудник")
                    continue
                file_nm_full = str(
                    'Output_files/HardSkills/' + str(curr_employee) + '.xlsx')
                # Создание файла в случае его отсутствия
                if not os.path.exists(file_nm_full):
                    print("Создание файла ", file_nm_full)
                    # Создаем файлы на основе пустого ДФ в 'Output_files/HardSkills/'
                    df_new.to_excel(file_nm_full)
                # Оценка самого себя заносится в свою же колонку своего файла
                if current_file_name.upper() == str(curr_employee).upper():
                    df_to_update = pd.read_excel(file_nm_full)
                    df_to_update.insert(2, curr_employee, df_assessment_hard.iloc[:, hard_iter])
                    # Удаление колонки с индексом
                    df_to_update = df_to_update.loc[:, ~df_to_update.columns.str.contains('^Unnamed')]
                    df_to_update.to_excel(file_nm_full)

                # Оцениваешь коллегу, значит надо занести колонку в файл коллеги с именем оценивающего
                else:
                    # поиск файла коллеги
                    if not os.path.exists(file_nm_full):
                        print('Файл отсутствует - ', )
                    df_to_update = pd.read_excel(file_nm_full)
                    # внесение оценки в файл коллеги (Бунаков оценил - в графу "Бунаков" файла "Бунцикин" помещается оценка)
                    df_to_update.insert(2, current_file_name, df_assessment_hard.iloc[:, hard_iter])
                    # Удаление колонки с индексом
                    df_to_update = df_to_update.loc[:, ~df_to_update.columns.str.contains('^Unnamed')]
                    # Перезаписываем эксельник с добавленной колонкой
                    df_to_update.to_excel(file_nm_full)

            hard_skills_list = list(df_assessment_hard.iloc[:, 0])
            print('End hard for ', current_file_name)

            # достать информацию по оценке своей и коллег-разработчиков SOFT
            df_assessment_soft = pd.read_excel(file
                                               , nrows=15
                                               , skiprows=42
                                               , sheet_name='Оценка разработчиков'
                                               , usecols=range(2, 23))
            df_assessment_soft['Employee'] = current_file_name
            df_assessment_soft_general.append(df_assessment_soft)
            df_new = pd.DataFrame(index=range(0, 14))

            df_new.insert(0, "Критерий\Сотрудник", df_assessment_soft["Критерий\Сотрудник"])
            print('Start Soft for ', current_file_name)
            j = -1
            for current_column in df_assessment_soft.columns:
                if current_column != 'Критерий\Сотрудник':

                    j = j + 1
                    curr_employee = str(list(df_assessment_soft.columns)[j])
                    # Если разбираемая колонка = 'Критерий\Сотрудник', то пропускаем итерацию
                    if str(curr_employee).upper() == str('Критерий\Сотрудник').upper():
                        print("Пропускаем итерацию для колонки Критерий\Сотрудник")
                        continue

                    file_nm_full = str(
                        'Output_files/SoftSkills/' + str(curr_employee) + '.xlsx')

                    if not os.path.exists(file_nm_full):
                        print("Создание файла ", file_nm_full)
                        # Создаем файлы на основе пустого ДФ в 'Output_files/HardSkills/'
                        df_new.to_excel(file_nm_full)

                    # Оценка самого себя заносится в свою же колонку своего файла
                    if current_file_name.upper() in str(curr_employee).upper():
                        df_to_update = pd.read_excel(file_nm_full)
                        df_to_update.insert(2, curr_employee, df_assessment_soft.iloc[:, j])
                        # Удаление колонки с индексом
                        df_to_update = df_to_update.loc[:, ~df_to_update.columns.str.contains('^Unnamed')]
                        df_to_update.to_excel(file_nm_full)

                    # Оцениваешь коллегу, значит надо занести колонку в файл коллеги с именем оценивающего
                    else:
                        # поиск файла коллеги
                        df_to_update = pd.read_excel(file_nm_full)
                        # внесение оценки в файл коллеги (Бунаков оценил - в графу "Бунаков" файла "Бунцикин" помещается оценка)
                        df_to_update.insert(2, current_file_name, df_assessment_soft.iloc[:, j])
                        # Удаление колонки с индексом
                        df_to_update = df_to_update.loc[:, ~df_to_update.columns.str.contains('^Unnamed')]
                        # Перезаписываем эксельник с добавленной колонкой
                        df_to_update.to_excel(file_nm_full)
            print('End Soft for ', current_file_name)

            # инфмормация с листа Cross-assessment
            df_cross = pd.read_excel(file
                                     , nrows=12
                                     , skiprows=7
                                     , sheet_name='Cross Разработчик-Аналитик '
                                     , usecols=range(2, 29))
            df_new = pd.DataFrame(index=range(0, 11))
            df_new.insert(0, "Критерий\Сотрудник", df_cross["Критерий\Сотрудник"])
            print('Start Cross for ', current_file_name)
            cross_iter = -1
            for current_column in df_cross.columns:
                if current_column != 'Критерий\Сотрудник':
                    cross_iter = cross_iter + 1
                    curr_employee = str(list(df_cross.columns)[cross_iter])
                    # Если разбираемая колонка = 'Критерий\Сотрудник', то пропускаем итерацию
                    if str(curr_employee).upper() == str('Критерий\Сотрудник').upper():
                        print("Пропускаем итерацию для колонки Критерий\Сотрудник")
                        continue
                    file_nm_full = str(
                        'Output_files/Cross/' + str(curr_employee) + '.' + str(file.name.split('.')[1]))
                    print('Разбираем оценки из файла - ', current_file_name, ' ;', 'Значение текущей колонки: ',
                          curr_employee, ' :')

                    if i == 1:
                        # Создаем файлы на основе пустого ДФ
                        df_new.to_excel(file_nm_full)

                    # Оцениваешь коллегу, значит надо занести колонку в файл коллеги с именем оценивающего
                    print('Cross OTHER')
                    # поиск файла коллеги-аналитика (или разработчика)
                    df_to_update = pd.read_excel(file_nm_full)
                    # внесение оценки в файл коллеги (Бунаков оценил - в графу "Бунаков" файла "Бунцикин" помещается оценка)
                    df_to_update.insert(2, current_file_name, df_cross.iloc[:, cross_iter])
                    # Удаление колонки с индексом
                    df_to_update = df_to_update.loc[:, ~df_to_update.columns.str.contains('^Unnamed')]
                    # Перезаписываем эксельник с добавленной колонкой
                    df_to_update.to_excel(file_nm_full)
            print('End Cross for ', current_file_name)

    # pd.set_option("display.width", 100)
    # df_qna_general_res = pd.DataFrame(pd.concat(df_qna_general))
    # df_qna_general_res.to_excel('Output_files/df_qna_general_res.xlsx')
    # df_assessment_hard_general_res = pd.concat(df_assessment_hard_general)
    # df_assessment_hard_general_res.to_excel('Output_files/df_assessment_hard_general_res.xlsx')
    # df_assessment_soft_general_res = pd.concat(df_assessment_soft_general)
    # df_assessment_soft_general_res.to_excel('Output_files/df_assessment_soft_general_res.xlsx')
    # df_cross_general_res = pd.concat(df_cross_general)
    # df_cross_general_res.to_excel('Output_files/df_cross_general_res.xlsx')

    # Сборка в единый DF по QNA
    qna_transposed_general = pd.concat(qna_transposed_general_pr)
    # print(qna_transposed_general)
    qna_transposed_general.to_excel('Output_files/result.xlsx')

    # print(qna_transposed_general.loc['Билибин'])


