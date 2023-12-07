import pandas as pd

from main import clear_folder

if __name__ == '__main__':
    employee_num = 0
    # for file in Path('Input_files').iterdir():
    #     if file.name != '.DS_Store' and file.name != '':
    clear_folder('Output_files/Res_Hard_soft')
    # переписать на захват первого файла в Output_files/HardSkills/
    file = 'Output_files/HardSkills/Ахтамьянова Гульназ Ринатовна.xlsx'
    df_for_list_employee= pd.read_excel(file
                           # , nrows=18
                       # , skiprows=15
                       # , sheet_name='Оценка разработчиков'
                       # , usecols=range(2, 23)
                       )


    df_for_list_employee = df_for_list_employee.loc[:, ~df_for_list_employee.columns.str.contains('^Unnamed')]
    df_for_list_employee = df_for_list_employee.loc[:, ~df_for_list_employee.columns.str.contains('^Критерий')]
    list_of_columns = list(df_for_list_employee.columns)
    print(list_of_columns)
    #сюда вставить цикл по списку колонок (сотрудников)
    for current_employee in list_of_columns:
        #вычитываем сначал хард файл для сотрудника
        df_current_employee_hard = pd.read_excel(str('Output_files/HardSkills/' + current_employee + '.xlsx')
                                                 )
        df_current_employee_hard = df_current_employee_hard.loc[:, ~df_current_employee_hard.columns.str.contains('^Unnamed')]
        # #затем читаем софт
        df_current_employee_soft = pd.read_excel(str('Output_files/SoftSkills/' + current_employee + '.xlsx')
                                                 )
        df_current_employee_soft = df_current_employee_soft.loc[:, ~df_current_employee_soft.columns.str.contains('^Unnamed')]
        # # объединяем файлы
        df_result = pd.concat([df_current_employee_hard, df_current_employee_soft])
        df_result.to_excel(str('Output_files/Res_Hard_Soft/' + current_employee + '.xlsx'))
