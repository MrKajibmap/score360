from pathlib import Path

import numpy as np
import pandas as pd

if __name__ == '__main__':
    # general_score_stat_df = pd.read_excel('Output_files/SoftSkills/Ахтамьянова Гульназ Ринатовна.xlsx')
    # general_score_stat_df = general_score_stat_df.loc[:,
    #                                 general_score_stat_df.columns.str.contains('^Крит')]
    # general_score_stat_df = pd.DataFrame(columns= ['Employee', 'Total_Score',
    # 'Total_Score_Cnt', 'Hard_Avg_Score', 'Hard_Scores_Cnt',
    # 'Soft_Avg_Score', 'Soft_Scores_Cnt'])
    # print(general_score_stat_df)
    main_iter = 0
    for file in Path('Output_files/Res_Hard_Soft').iterdir():
        if file.name != '.DS_Store' and file.name != '':
            main_iter = main_iter + 1
            print(file)
            current_employee = str(file.name).split(sep='.')[0]
            df_current_employee_hard = pd.read_excel(file)
            df_current_employee_hard = df_current_employee_hard.loc[:,
                                       ~df_current_employee_hard.columns.str.contains('^Unnamed')]
            # убираем владельца файла, чтобы не учитывать его в средней оценке
            df_current_employee_hard = df_current_employee_hard.loc[:,
                                       ~df_current_employee_hard.columns.str.contains(str('^' + current_employee))]
            count_row = df_current_employee_hard.shape[0]  # Gives number of rows
            count_col = df_current_employee_hard.shape[1]  # Gives number of columns
            list_cnt_not_null_val = []
            list_col_sum = []
            list_col_avg = []
            for row in range(0, count_row):
                col_sum = 0
                col_count_not_Null_values = 0
                col_avg = 0
                for col in range(1, count_col):
                    if str(df_current_employee_hard.iloc[row, col]).upper() == "Б/О" or df_current_employee_hard.iloc[
                        row, col] is None or str(df_current_employee_hard.iloc[row, col]).upper() == "NAN" or str(
                        df_current_employee_hard.iloc[row, col]).upper() in ('-', '.'):
                        # col_count_not_Null_values = col_count_not_Null_values
                        # col_sum = col_sum
                        # col_avg = col_avg
                        pass
                    else:
                        col_count_not_Null_values = col_count_not_Null_values + 1
                        col_sum = col_sum + int(df_current_employee_hard.iloc[row, col])
                        col_avg = col_sum / col_count_not_Null_values
                list_cnt_not_null_val.append(col_count_not_Null_values)
                list_col_sum.append(col_sum)
                list_col_avg.append(col_avg)
            df_current_employee_hard['Count'] = list_cnt_not_null_val
            df_current_employee_hard['Sum'] = list_col_sum
            df_current_employee_hard['Avg'] = list_col_avg
            df_current_employee_hard.to_excel(str('Output_files/Results_with_stats/' + file.name))
            #     #
            df_current_employee_with_stat = pd.read_excel(str('Output_files/Results_with_stats/' + file.name))
            df_current_employee_with_stat = df_current_employee_with_stat.loc[:,
                                            ~df_current_employee_with_stat.columns.str.contains('^Unnamed')]
            count_row = df_current_employee_with_stat.shape[0]  # Gives number of rows
            count_col = df_current_employee_with_stat.shape[1]  # Gives number of columns

            df_current_empl = pd.DataFrame(data=df_current_employee_with_stat.loc[1,
                                                                                  df_current_employee_with_stat.columns.str.contains(
                                                                                      '^Крит')])
            df_current_empl['Employee'] = current_employee
            df_current_empl['Total_Score'] = round(df_current_employee_with_stat.iloc[:, count_col - 1].sum(), 1)
            df_current_empl['Total_Score_Cnt'] = round(df_current_employee_with_stat.iloc[:, count_col - 3].mean(), 1)
            df_current_empl['Hard_Avg_Score'] = round(df_current_employee_with_stat.iloc[0:17, count_col - 1].sum(), 1)
            df_current_empl['Hard_Scores_Cnt'] = round(df_current_employee_with_stat.iloc[0:17, count_col - 3].mean(),
                                                       1)
            df_current_empl['Soft_Avg_Score'] = round(df_current_employee_with_stat.iloc[17:31, count_col - 1].sum(), 1)
            df_current_empl['Soft_Scores_Cnt'] = round(df_current_employee_with_stat.iloc[17:31, count_col - 3].mean(),
                                                       1)
            new_df = pd.read_excel('Output_files/Res_Hard_Soft/' + file.name)
            new_df = new_df.loc[:,
                     new_df.columns.str.contains(str('^' + str(file.name).split(sep='.')[0]))]
            df_current_empl['Self_Hard'] = new_df.iloc[0:17, 0].sum()
            df_current_empl['Self_Soft'] = new_df.iloc[17:31, 0].sum()
            df_current_empl['Self_Total'] = new_df.iloc[0:31, 0].sum()
            df_current_empl.to_excel(str('Output_files/Self_Team_ws_Stat/' + file.name))
            # print(df_current_empl)
            if main_iter == 1:
                # для первой итерации  создаем файл с первым сотрудником
                general_score_stat_df = df_current_empl
                # print(general_score_stat_df)
                # print('PRIVET')
            else:
                # накапливаем статистику по каждому сотруднику (проходясь по файлам) в итоговый ДФ
                general_score_stat_df = pd.concat([general_score_stat_df, df_current_empl])
                # print(general_score_stat_df)
            # дообогащение файла селф оценкой
            df_current_employee_with_stat_added = pd.read_excel(str('Output_files/Results_with_stats/' + file.name))
            # df_current_employee_with_stat_added = pd.DataFrame(data=df_current_employee_with_stat_added.loc[:,
            #                                                         ~df_current_employee_with_stat_added.columns.str.contains(
            #                                                             '^Крит')])
            df_current_employee_with_stat_added = pd.DataFrame(data=df_current_employee_with_stat_added.loc[:,
                                                                    ~df_current_employee_with_stat_added.columns.str.contains(
                                                                        '^Unnamed')])
            # забираем самооценку
            df_current_employee_self_assessment = pd.read_excel(str('Output_files/Res_Hard_Soft/' + file.name))
            df_current_employee_self_assessment = pd.DataFrame(data=df_current_employee_self_assessment.loc[:,
                                                                    df_current_employee_self_assessment.columns.str.contains(
                                                                        current_employee)])
            # добавляем ее к оценкам от коллег + статистика
            df_current_employee_with_stat_added.insert(1, current_employee, df_current_employee_self_assessment.loc[:,
                                                                            df_current_employee_self_assessment.columns.str.contains(
                                                                                current_employee)])

            df_current_employee_with_stat_added.to_excel(str('Output_files/Self_Team_ws_Stat/' + file.name))

            # df_current_employee_self_assessment = pd.DataFrame(data=df_current_employee_self_assessment.loc[1, df_current_employee_self_assessment.columns.str.contains()
            # добавление на отдельные странички всех сотрудников в финальный файл

        # general_score_stat_df = pd.DataFrame(data=general_score_stat_df.loc[:,
        #                                           ~general_score_stat_df.columns.str.contains(
        #                                               str('1'))])
        general_score_stat_df.to_excel(str('Output_files/Results/dev_scores_results.xlsx'))


    for file in Path('Output_files/Results_with_stats').iterdir():
        if file.name != '.DS_Store' and file.name != '':
            current_employee = str(file.name).split(sep='.')[0]
            df_current_employee_with_stat = pd.read_excel(file)
            df_current_employee_with_stat = df_current_employee_with_stat.loc[:,
                                            ~df_current_employee_with_stat.columns.str.contains('^Unnamed')]
            print('>> << >> << ' , current_employee)
            with pd.ExcelWriter(
                    'Output_files/Results/dev_scores_results.xlsx',
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="new",
            ) as writer:
                df_current_employee_with_stat.to_excel(writer, sheet_name=current_employee)
