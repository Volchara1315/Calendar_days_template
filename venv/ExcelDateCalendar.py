import xlsxwriter
from xlsxwriter.utility import xl_range


def sav_l_cal_excel(list_month_day, n_year, str_path_file):

    i = 0
    j = 0
    i_month = 0
    i_day = 0
    str_buf_day = ''
    str_buf_month = ''
    cell_range = ''

    l_worksheet = ['01 (Січень)', '02 (Лютий)', '03 (Березень)', '04 (Квітень)',
                   '05 (Травень)', '06 (Червень)', '07 (Липень)', '08 (Серпень)',
                   '09 (Вересень)', '10 (Жовтень)', '11 (Листопад)', '12 (Грудень)']

    l_size_col = [15, 15, 25, 25, 10,
                     12, 12, 20, 20, 40]

    l_tit = ['Дата', 'День тижня', 'Робочий графік', 'Обід', 'К-ть. год.',
             'Сума(грн)', 'Робочі дні', 'Проїзд до роботи', 'Проїзд до додому', 'Опис дня']

    l_basic_buf = ['', '', '00:00:00 - 00:00:00', '00:00:00 - 00:00:00',
                 0, 0, 0, 0, 0, '']

    l_buf_form = ['', '', '', '', 0, 0, 0, 0, 0, '']

    workbook = xlsxwriter.Workbook(str_path_file)

    title_format = workbook.add_format({'bold': True, 'font_color': 'black', 'bg_color': 'white',
                                        'font_size': 12, 'align': 'center', 'border': 5,
                                        'border_color': 'black'})

    weekend_format = workbook.add_format({'font_color': 'black', 'bg_color': '#FF0000',
                                          'font_size': 12, 'align': 'center', 'border': 5,
                                          'border_color': 'black'})

    buf_format = workbook.add_format({'font_color': 'black', 'bg_color': '#757171',
                                      'font_size': 12, 'align': 'center', 'border': 5,
                                      'border_color': 'black'})

    basic_format = workbook.add_format({'font_color': 'black', 'bg_color': 'white',
                                        'font_size': 12, 'align': 'center', 'border': 5,
                                        'border_color': 'black'})

    while i_month <= 11:
        worksheet = workbook.add_worksheet(l_worksheet[i_month])

        for x_col in l_size_col:
            worksheet.set_column(i, i, x_col)
            i = i + 1
        i = 0

        for x_tit in l_tit:
            worksheet.write(0, i, x_tit, title_format)
            i = i + 1
        i = 1

        while i_day < len(list_month_day[i_month]):
            if list_month_day[i_month][i_day] == 'Нд':
                l_basic_buf[9] = 'Вихідний'
                while j < 10:
                    if j == 0:
                        if i_day < 9:
                            str_buf_day = f'0{i_day + 1}'
                        else:
                            str_buf_day = f'{i_day + 1}'
                        if i_month < 9:
                            str_buf_month = f'0{i_month + 1}'
                        else:
                            str_buf_month = f'{i_month + 1}'
                        worksheet.write(i, j, str(f'{str_buf_day}.{str_buf_month}.{n_year}'),
                                        weekend_format)
                    elif j == 1:
                        worksheet.write(i, j, list_month_day[i_month][i_day], weekend_format)
                    else:
                        worksheet.write(i, j, l_basic_buf[j], weekend_format)
                    j = j + 1
                i = i + 1
                j = 0

                for x in l_buf_form:
                    worksheet.write(i, j, x, buf_format)
                    j = j + 1
                j = 0
            else:
                l_basic_buf[9] = ''
                while j < 10:
                    if j == 0:
                        if i_day < 9:
                            str_buf_day = f'0{i_day + 1}'
                        else:
                            str_buf_day = f'{i_day + 1}'
                        if i_month < 9:
                            str_buf_month = f'0{i_month + 1}'
                        else:
                            str_buf_month = f'{i_month + 1}'
                        worksheet.write(i, j, str(f'{str_buf_day}.{str_buf_month}.{n_year}'),
                                        basic_format)
                    elif j == 1:
                        worksheet.write(i, j, list_month_day[i_month][i_day], basic_format)
                    else:
                        worksheet.write(i, j, l_basic_buf[j], basic_format)
                    j = j + 1
                j = 0

            i = i + 1
            i_day = i_day + 1

        for x_f in l_buf_form:
            #worksheet.write(i, j, x_f, title_format)
            if j > 3:
                cell_range = xl_range(1, j, i, j)
                worksheet.write(i, j, f'=SUM({cell_range})', title_format)
            j = j + 1
        j = 0
        worksheet.merge_range(i, 0, i, 3, 'Результат за місяць - 0 грн', title_format)

        i_month = i_month + 1
        i = 0
        i_day = 0

    workbook.close()
    return True