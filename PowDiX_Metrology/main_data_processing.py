"""
Main logical part:
- import data from 'excel';
- reading and clearing data;
- metrology calculation;
- converting calculations to the same view for postprocessing.
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

# to change the number of lines read from the document, change this value
pars_zone_numbers = 20
# to change the number of columns read from the document, change this value
pars_zone_len = 25
# list of column-rows to read data
pars_list = []

directory_expand = None

srm_name = 'not read'
number_of_diffractograms = 0

main_window_destroy = True
# trigger to recreate the main window when you run the program again, if no data was found
check_for_no_answer = False


def pars_location(i=pars_zone_len):
    """Generates a list of coordinates to read the data from an Excel document"""

    global pars_list
    pars_list = []
    pars_zone_len_converted = int(i / 25)

    # creates a list of letters by the value of the pars_zone_len variable and generates a letter-by-letter parses
    # area based on it
    pars_zone_len_sheet = []
    pars_zone_letters = ['']

    for x in [chr(i) for i in range(ord('A'), ord('Z') + 1)]:
        pars_zone_len_sheet.append(x)
    for x in range(pars_zone_len_converted):
        pars_zone_letters.append(pars_zone_len_sheet[x])

    # fills in a parsed list to enumerate the columns of the document with letters from A to Z, AA to AZ, etc
    for zone in pars_zone_letters:
        for x in [chr(i) for i in range(ord('A'), ord('Z') + 1)]:
            pars_list.append(str(zone) + x)


def file_location(directory):
    """Opens the file selection menu and loads the selected Excel"""

    global sheet
    global directory_expand

    if not directory:
        filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
        if filename == '':
            return True
        else:
            directory_expand = filename

            # loads an Excel table into the openpyxl library
            wb = openpyxl.load_workbook(filename=filename)
            sheet = wb.active
    else:
        wb = openpyxl.load_workbook(filename=directory)
        sheet = wb.active


def pars(range_for_pars=pars_zone_numbers):
    """Searches all the necessary data in the uploaded sheet, cleans it, and converts it into a convenient form"""

    global nist_peak_position
    global data_position
    global pars_zone_numbers
    global pars_zone_len
    global main_window_destroy
    global check_for_no_answer
    global number_of_diffractograms
    global srm_name
    number_of_diffractograms = 0
    data_position = {}

    def clear_unnecessary_elements(read_values, first_index_to_clear):
        """Removes unnecessary data from the unloaded list, leaving only intensities, hkl, and peak positions"""

        if None in read_values:
            del read_values[0:first_index_to_clear + 1]
            last_index_none = read_values.index(None)
            del read_values[last_index_none:]

    for trigger in pars_list:

        # reads each column in the document in turn within the bounds
        # (pars_zone_len_numbers by columns; pars_zone_len by rows) and loads it into the values variable
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

            elif 'SRM 1976b' in values:
                first_index = values.index('SRM 1976b')
                clear_unnecessary_elements(values, first_index)
                nist_peak_position = values
                srm_name = 'SRM 1976b'

    # check for read data
    if len(data_position) < 2:
        res = messagebox.askquestion('Данные не найдены', 'Расширить зону поиска в документе?')

        # if "yes", expands the search area and starts the process anew
        if res == 'yes':
            main_window_destroy = False
            pars_zone_numbers += 60
            pars_zone_len += 75
            pars_location(pars_zone_len)
            file_location(directory_expand)
            pars(pars_zone_numbers)
        elif res == 'no':
            # cancel trigger, because otherwise there are problems with the main window
            check_for_no_answer = True
            return None
        else:
            messagebox.showwarning('error', 'Something went wrong!')


def sko_init(error, peak_hkl_1):
    """The function of pre-processing data before the staple (creating dictionaries with the right number of items)"""

    global index_list
    global error_calculated_intensities
    global error_calculated_positions
    sko_result = {}

    def sko(intensities_positions):
        """Finds the standard deviations of the intensity / position"""

        summ = 0
        summ_for_sqrt = 0
        squared_difference = []
        for data_for_sko in intensities_positions:
            if data_for_sko != 0 and len(intensities_positions) != 0:
                summ += float(data_for_sko)
        average = summ / len(intensities_positions)
        # squared difference
        for data_for_sko in intensities_positions:
            difference = (float(data_for_sko) - average) * (float(data_for_sko) - average)
            squared_difference.append(difference)
        # the root of the sum divided by the number of elements minus 1
        for items in squared_difference:
            summ_for_sqrt += items
        sqroot = math.sqrt(summ_for_sqrt / (len(squared_difference) - 1))
        # the result of the previous step multiplied by 100 and divided by the average value for the intensities
        sko_results = sqroot * 100 / average
        result = {'exp_data': intensities_positions,
                  'average': average,
                  'sqroot': sqroot,
                  'sko': sko_results,
                  'metrology_var': 1,
                  }
        return result

    # generates two arrays for data on the length of read peak values
    index_list = [str(i) for i in range(0, len(peak_hkl_1))]
    for i in index_list:
        error_calculated_positions.update({i: []})
    for i in index_list:
        error_calculated_intensities.update({i: []})

    # conversion of intensity/position lists
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
    """Calculates the absolute error"""

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

    # reduction to a common view with the standard deviation of the intensities/positions
    for i in range(len(peak_hkl)):
        if 'None information' not in absolute_error_before_unification[str(i)] and absolute_error_before_unification[
                    str(i)] != []:
            absolute_error[peak_hkl[i]] = {'exp_data': absolute_error_before_unification[str(i)],
                                           'metrology_var': 1,
                                           }
        else:
            absolute_error[peak_hkl[i]] = 'None information'
