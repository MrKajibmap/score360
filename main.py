import os
import shutil

import numpy as np
import pandas as pd
from pathlib import Path

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
    for filename in os.listdir(Path('Output_files/HardSkills')):
        file_path = os.path.join(Path('Output_files/HardSkills'), filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Ошибка при удалении - %s. Причина: %s' % (file_path, e))

    # # Создание новых файлов для хард скиллов
    # for file in Path('Input_files').iterdir():
    #     shutil.copyfile(
    #         'Templates/Assessment_template.xlsx', os.path.join(Path('Output_files/HardSkills'), file.name))

    files = os.listdir(Path('Input_files'))
    print(files)

    i = 0
    for file in Path('Input_files').iterdir():
        if file.name != '.DS_Store' and file.name !='':
            i = i + 1
            current_file_name = file.name.split('.')[0]
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
            print('Start hard')
            # достать информацию по оценке своей и коллег-разработчиков HARD
            df_assessment_hard = pd.read_excel(file
                                               , nrows=18
                                               , skiprows=15
                                               , sheet_name='Оценка разработчиков'
                                               , usecols=range(2, 23))
            list_of_columns = list(df_assessment_hard.columns)
            # del_first_clmn = list_of_columns.pop(0)
            df_new = pd.DataFrame(index=range(0, 17))
            # print(df_new.columns)

            # extracted_col = df_assessment_hard["Критерий\Сотрудник"]
            # print(extracted_col)
            df_new.insert(0, "Критерий\Сотрудник", df_assessment_hard["Критерий\Сотрудник"])
            # print(df_new.columns)
            # df_assessment_hard['Employee'] = current_file_name
            # df_assessment_hard_general.append(df_assessment_hard)
            j = -1
            # находясь в файле, разбираем оценки как самого себя, так и коллег, пройдя по всем оценкам всех коллег (колонки = коллеги)
            for current_column in df_assessment_hard.columns:
                if current_column != 'Критерий\Сотрудник':
                    j = j +1
                    curr_employee = str(list(df_assessment_hard.columns)[j])
                    # str(list(df_assessment_hard.columns)[j]).split()[0]
                    file_nm_full = str('Output_files/HardSkills/' + str(curr_employee).split()[0] + '.' + str(file.name.split('.')[1]))
                    print('Разбираем оценки из файла - ', current_file_name, ' ;', 'Значение текущей колонки: ' , curr_employee ,' :')
                    print(str(list(df_assessment_hard.columns)[j]).split()[0])
                    print(df_assessment_hard.iloc[:,j])

                    if i == 1:
                        # Создаем файлы на основе пустого ДФ
                        df_new.to_excel(file_nm_full)
                    # Оценка самого себя заносится в свою же колонку своего файла
                    if current_file_name.upper() in str(curr_employee).upper():
                        df_to_update = pd.read_excel(file_nm_full)
                        print('first:', df_to_update.columns)
                        print('second:', df_to_update.columns)
                        df_to_update.insert(2, curr_employee,  df_assessment_hard.iloc[:, j])
                        df_to_update.to_excel(file_nm_full)
                    # Оценка коллеги заносится в его файл, но с ренеймом/
                    # Оцениваешь коллегу, значит надо занести колонку в файл коллеги с именем оценивающего

                    else:
                        # поиск файла коллеги
                        df_to_update = pd.read_excel(file_nm_full)
                        # внесение оценки в файл коллеги (Бунаков оценил - в графу "Бунаков" файла "Бунцикин" помещается оценка)
                        df_to_update.insert(2, current_file_name, df_assessment_hard.iloc[:, j])
                        # Перезаписываем эксельник с добавленной колонкой
                        df_to_update.to_excel(file_nm_full)
                        df_assessment_hard.iloc[:, j]
            # print(df_assessment_hard.iloc[:, 1])
            hard_skills_list = list(df_assessment_hard.iloc[:, 0])

            # print('End hard')

            # достать информацию по оценке своей и коллег-разработчиков SOFT
            # df_assessment_soft = pd.read_excel(file
            #                                    , nrows=15
            #                                    , skiprows=42
            #                                    , sheet_name='Оценка разработчиков'
            #                                    , usecols=range(2, 23))
            # df_assessment_soft['Employee'] = current_file_name
            # df_assessment_soft_general.append(df_assessment_soft)

            # инфмормация с листа Cross-assessment
            # df_cross = pd.read_excel(file
            #                          , nrows=12
            #                          , skiprows=7
            #                          , sheet_name='Cross Разработчик-Аналитик '
            #                          , usecols=range(2, 29))
            # df_cross['Employee'] = current_file_name
            # df_cross_general.append(df_cross)

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