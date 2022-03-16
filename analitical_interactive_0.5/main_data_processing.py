"""
Main logical part:
- import data from 'excel';
- reading and clearing data;
- metrology calculation;
- converting calculations to the same view for postprocessing;
"""

import math
from tkinter.filedialog import askopenfilename
import openpyxl
from tkinter import messagebox

# variable for raw data from 'excel'
sheet = []

# imported list of peak hkl
peak_hkl = []

index_list = []
nist_peak_position = []

error_calculated_intensities = {}
error_calculated_positions = {}

data_intensity = {}
data_position = {}

absolute_error = {}

# для изменения количества считываемых из документа строк, изменить эту величину
pars_zone_numbers = 20
# для изменения количества считываемых из документа столбцов, изменить эту величину
pars_zone_len = 25
# список столбцов-строк для считывания данных
pars_list = []

directory_expand = None

srm_name = 'not readen'
number_of_diffractograms = 0

main_window_destroy = True
# триггер для пересоздания главного окна при повторной прогонке программы, если данные не были найдены
check_for_no_answer = False



def pars_location(i=pars_zone_len):
    """Генерирует список координат для последующего считывания данных из документа excel"""
    global pars_list
    pars_list = []
    pars_zone_len_converted = int(i / 25)
    # по значению переменной pars_zone_len создает список из букв и на его основе генеирует область парса по буквам
    pars_zone_len_sheet = []
    pars_zone_letters = ['']
    for x in [chr(i) for i in range(ord('A'), ord('Z') + 1)]:
        pars_zone_len_sheet.append(x)
    for x in range(pars_zone_len_converted):
        pars_zone_letters.append(pars_zone_len_sheet[x])
    # Заполняет парс-лист для перебора столбцов документа букварём от A до Z, от AA до AZ и т.д.:
    for zone in pars_zone_letters:
        for x in [chr(i) for i in range(ord('A'), ord('Z') + 1)]:
            pars_list.append(str(zone) + x)


def file_location(directory):
    """Открывает меню выбора файла и вгружает выбранный excel"""
    global sheet
    global directory_expand

    # Эта штука должна что-то скрывать, но нифига не делает
    # Tk().withdraw()

    if not directory:
        filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
        if filename == '':
            return True
        else:
            directory_expand = filename
            # вгружает таблицу excel в библиотеку openpyxl
            wb = openpyxl.load_workbook(filename=filename)
            sheet = wb.active
    else:
        # вгружает таблицу excel в библиотеку openpyxl
        wb = openpyxl.load_workbook(filename=directory)
        sheet = wb.active


def pars(range_for_pars=pars_zone_numbers):
    """Ищет в выгруженном листе все нужные данные, считывает, чистит, преобразует в удобный вид"""
    global nist_peak_position
    global data_position
    global pars_zone_numbers
    global pars_zone_len
    global main_window_destroy
    global check_for_no_answer
    global number_of_diffractograms
    global srm_name

    def clear_unnecessary_elements(readen_values, first_index_to_clear):
        """Удаляет ненужные данные из выгруженного списка, оставляя только интенсивности, hkl и положения пиков"""

        if None in readen_values:
            del readen_values[0:first_index_to_clear + 1]
            last_index_none = readen_values.index(None)
            del readen_values[last_index_none:]

    number_of_diffractograms = 0
    for trigger in pars_list:
        # по очереди считывает каждый столбец в документе в границах
        # (pars_zone_len_numbers по столбцам; pars_zone_len по строкам) и вгружает в переменную values
        values_variable = str(trigger) + '0:' + str(trigger) + '100'
        values = [v[0].value for v in sheet[values_variable]]

        for number in range(range_for_pars):
            if 'Spectrum ' + str(number) + ', 2θ' in values:
                clear_unnecessary_elements(values, values.index('Spectrum ' + str(number) + ', 2θ'))
                data_position['spectrum_' + str(number) + '_position'] = values
            if 'Spectrum ' + str(number) + ', counts' in values:
                number_of_diffractograms += 1
                clear_unnecessary_elements(values, values.index('Spectrum ' + str(number) + ', counts'))
                data_intensity['spectrum_' + str(number) + '_intensity'] = values

            if 'Peak' in values:
                first_index = values.index('Peak')
                clear_unnecessary_elements(values, first_index)
                peak_hkl_format = values
                peak_hkl.clear()
                for hkl in peak_hkl_format:
                    hkl = str(hkl).replace('-', '')
                    if hkl == '12':
                        hkl = '(012)'
                    elif hkl == '24':
                        hkl = '(024)'
                    elif '(' not in hkl:
                        hkl = '(' + hkl + ')'
                    peak_hkl.append(hkl)

            elif 'SRM 1976c' in values:
                first_index = values.index('SRM 1976c')
                clear_unnecessary_elements(values, first_index)
                nist_peak_position = values
                srm_name = 'SRM 1976c'
                print(values)

            elif 'SRM 1976b' in values:
                first_index = values.index('SRM 1976b')
                clear_unnecessary_elements(values, first_index)
                nist_peak_position = values
                srm_name = 'SRM 1976b'
                print(values)

    # проверка на наличие считанных данных
    if len(data_position) < 2:
        res = messagebox.askquestion('Data not found', 'Expand the search area?')

        # если "да", расширяет зону поиска и запускает процесс заново
        if res == 'yes':
            main_window_destroy = False
            pars_zone_numbers += 60
            pars_zone_len += 75
            pars_location(pars_zone_len)
            file_location(directory_expand)
            pars(pars_zone_numbers)
        elif res == 'no':
            # триггер отмены, ибо иначе возникают проблемы с главным окном
            check_for_no_answer = True
            return None
        else:
            messagebox.showwarning('error', 'Something went wrong!')


def sko_init(error, peak_hkl_1):
    """Функция предварительной обработки данных перед ско (в начале создание словарей с нужным количеством элементов)"""
    global index_list
    global error_calculated_intensities
    global error_calculated_positions
    sko_result = {}

    def sko(intensities_positions):
        """Находит СКО интенсивности / положения"""

        summ = 0
        summ_for_sqrt = 0
        squared_difference = []
        for data_for_sko in intensities_positions:
            if data_for_sko != 0 and len(intensities_positions) != 0:
                summ += float(data_for_sko)
        average = summ / len(intensities_positions)
        # разница в квадрате
        for data_for_sko in intensities_positions:
            difference = (float(data_for_sko) - average) * (float(data_for_sko) - average)
            squared_difference.append(difference)
        # корень от суммы, деленный на количество элементов минус 1
        for items in squared_difference:
            summ_for_sqrt += items
        sqroot = math.sqrt(summ_for_sqrt / (len(squared_difference) - 1))
        # результат предыдущего шага, умноженный на 100 и деленный на среднее значение (СКО) для интенсивностей
        sko_results = sqroot * 100 / average
        result = {'exp_data': intensities_positions,
                  'average': average,
                  'sqroot': sqroot,
                  'sko': sko_results,
                  'metrology_var': 1,
                  }
        return result

    # генерирует два массива для данных по длине считанных значений пиков
    index_list = [str(i) for i in range(0, len(peak_hkl_1))]
    for i in index_list:
        error_calculated_positions.update({i: []})
    for i in index_list:
        error_calculated_intensities.update({i: []})

    # преобразование списков интенсивностей/положений
    for key, value in data_intensity.items():
        timer = 0
        for item in value:
            error_calculated_intensities[str(timer)].append(str(item))
            timer += 1

    for key, value in data_position.items():
        timer = 0
        for item in value:
            error_calculated_positions[str(timer)].append(str(item))
            timer += 1

    timer = 0
    for name, value in error.items():
        if 'Not found' in value or '' in value:
            sko_result[peak_hkl_1[timer]] = 'None information'
        else:
            if not error[str(timer)]:
                pass
            else:
                message = sko(error[str(timer)])
                sko_result[peak_hkl_1[timer]] = message
        timer += 1
    return sko_result


def abs_err():
    """Вычисляет абсолютную погрешность"""
    global absolute_error

    absolute_error_before_unification = {}
    for i in index_list:
        absolute_error_before_unification.update({i: []})

    for name, peak_list in data_position.items():
        initial_position = peak_list[0]
        for peak_position in peak_list:
            if peak_position == 'Not found' or peak_position == '':
                difference = 'None information'
            else:
                difference_experimental = (float(peak_position) - float(initial_position))
                difference_nist = (float(nist_peak_position[peak_list.index(peak_position)])
                                   - float(nist_peak_position[0]))
                difference = abs(difference_nist - difference_experimental)
            absolute_error_before_unification[str(peak_list.index(peak_position))].append(difference)

    # Приведение к общему виду с СКО интенсивности/положения
    for i in range(len(peak_hkl)):
        if 'None information' not in absolute_error_before_unification[str(i)] and absolute_error_before_unification[
                    str(i)] != []:
            absolute_error[peak_hkl[i]] = {'exp_data': absolute_error_before_unification[str(i)],
                                           'metrology_var': 1,
                                           }
        else:
            absolute_error[peak_hkl[i]] = 'None information'
