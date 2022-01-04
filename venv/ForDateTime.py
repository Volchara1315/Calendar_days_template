from datetime import datetime, date, time

def cr_l_cal(n_year):
    #global i
    i = 0
    count_day_year = 0
    n_month_day = 1
    list_count_day = [31, 28, 31, 30,
                      31, 30, 31, 31,
                      30, 31, 30, 31]
    list_str_weekday = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд']
    list_monthday_weekday = []

    if (n_year % 4 == 0):
        count_day_year = 366
        list_count_day[1] = 29
    else:
        count_day_year = 365
        list_count_day[1] = 28

    while i <= 11:
        list_monthday_weekday.append([])
        while n_month_day <= list_count_day[i]:
            list_monthday_weekday[i].append(
                list_str_weekday[datetime(n_year, int(i + 1), n_month_day, 0, 0, 0).weekday()])
            n_month_day = n_month_day + 1
        n_month_day = 1
        i = i + 1

    return list_monthday_weekday