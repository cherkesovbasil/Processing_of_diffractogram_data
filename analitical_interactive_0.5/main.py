"""
Secondary interface:
- data postprocessing;
- generating main GUI;
- generating report (in development);
"""

from main_data_processing import *
from tkinter import *

sko_result_intens = {}
sko_result_pos = {}

pars_zone_numbers_local = 20
pars_zone_len_local = 25

sko_pos = []
sko_int = []
abs_pos = []

something = main_window_destroy
reload_x = 0
secondary_window = None
secondary_window_destroy = False


def message_for_frames():
    """Calculates minimum, average and maximum error values for output to the main frame"""

    global sko_int
    global sko_pos
    global abs_pos
    global something
    global sko_result_intens
    global sko_result_pos

    # first - minimum, second - average, third - maximum error for each variable
    sko_pos = [[0, 1000], 0.0, [0, 0]]
    sko_int = [[0, 1000], 0.0, [0, 0]]
    abs_pos = [[0, 1000], 0.0, [0, 0]]

    # min\max\avr in standard deviation of intensities
    list_for_min_max_avr = []
    for pos, data in sko_result_intens.items():
        if data != 'None information' and sko_result_intens[pos]['metrology_var'] == 1:
            list_for_min_max_avr.append(sko_result_intens[pos]['sko'])
        else:
            list_for_min_max_avr.append(1964)
        compare = min(list_for_min_max_avr)
        compare_2 = max(list_for_min_max_avr)
        if compare < sko_int[0][1]:
            sko_int[0].clear()
            sko_int[0].append(pos)
            sko_int[0].append(compare)
        if compare_2 > sko_int[2][1] and compare_2 != 1964:
            sko_int[2].clear()
            sko_int[2].append(pos)
            sko_int[2].append(compare_2)
        while 1964 in list_for_min_max_avr:
            list_for_min_max_avr.remove(1964)
        for things in list_for_min_max_avr:
            sko_int[1] += things
        if len(list_for_min_max_avr) > 0:
            sko_int[1] = sko_int[1] / (len(list_for_min_max_avr) + 1)

    # min\max\avr in standard deviation of positions
    list_for_min_max_avr = []
    for pos, data in sko_result_pos.items():
        if data != 'None information' and sko_result_pos[pos]['metrology_var'] == 1 and data['exp_data']:
            list_for_min_max_avr.append(sko_result_pos[pos]['sqroot'])
        else:
            list_for_min_max_avr.append(1964)
        compare = min(list_for_min_max_avr)
        compare_2 = max(list_for_min_max_avr)
        if compare < sko_pos[0][1]:
            sko_pos[0].clear()
            sko_pos[0].append(pos)
            sko_pos[0].append(compare)
        if compare_2 > sko_pos[2][1] and compare_2 != 1964:
            sko_pos[2].clear()
            sko_pos[2].append(pos)
            sko_pos[2].append(compare_2)
        while 1964 in list_for_min_max_avr:
            list_for_min_max_avr.remove(1964)
        for things in list_for_min_max_avr:
            sko_pos[1] += things
        if len(list_for_min_max_avr) > 0:
            sko_pos[1] = sko_pos[1] / (len(list_for_min_max_avr) + 1)

    # min\max\avr in absolute error
    list_for_min_max_avr = []
    for pos, data in absolute_error.items():
        if 'None information' in data or not data['exp_data'] or data['metrology_var'] == 0:
            list_for_min_max_avr.append(0.000001)
        else:
            list_for_min_max_avr.append(max(data['exp_data']))
        abs_pos[2] = [peak_hkl[list_for_min_max_avr.index(max(list_for_min_max_avr))], max(list_for_min_max_avr)]

    list_for_min_max_avr = []
    for pos, data in absolute_error.items():
        if 'None information' not in data and data['metrology_var'] == 1 and pos != str(peak_hkl[0]):
            list_for_min_max_avr.append(min(data['exp_data']))
        else:
            list_for_min_max_avr.append(1964)
        abs_pos[0] = [peak_hkl[list_for_min_max_avr.index(min(list_for_min_max_avr))], min(list_for_min_max_avr)]
    summ = 0
    len_timer = 0
    for pos, data in absolute_error.items():
        if 'None information' not in data and 0.0 not in data and data['exp_data'] and data['metrology_var'] == 1:
            for items in data['exp_data']:
                summ += items
                len_timer += 1
    if len_timer == 0:
        abs_pos[1] = 0.0
    elif len(list_for_min_max_avr) > 0:
        abs_pos[1] = summ / len_timer

    # if only one element for processing left
    if abs_pos[0][1] == 1964:
        abs_pos[0][1] = abs_pos[2][1]

    something = True


def change_search_area():
    """Creates two sliders to select the number of rows and columns to be read from Excel file; writes it in global"""

    def write_pars_numbers(self):
        global pars_zone_numbers_local
        pars_zone_numbers_local = (pars_zone.get())
        print(self)

    def write_pars_len(self):
        global pars_zone_len_local
        pars_zone_len_local = (pars_len.get())
        print(self)

    search_area_button.destroy()
    pars_zone = IntVar()
    scale_numbers = Scale(frame1, from_=10, to=200, orient=HORIZONTAL, label='Количество строк: ',
                          resolution=10, length=140, variable=pars_zone,
                          command=write_pars_numbers)
    scale_numbers.set(20)
    scale_numbers.pack()

    pars_len = IntVar()
    scale_len = Scale(frame1, from_=25, to=250, orient=HORIZONTAL, label='Количество столбцов: ',
                      resolution=25, length=140, variable=pars_len,
                      command=write_pars_len)
    scale_len.set(50)
    scale_len.pack()


def open_file():
    """Executes the program body (calculation of deviations and errors)"""

    global sko_result_pos
    global sko_result_intens
    global reload_x
    global secondary_window
    global secondary_window_destroy

    if main_window_destroy:
        pars_location(pars_zone_len_local)
        trigger = file_location(directory_expand)
        if trigger:
            return None
    pars(pars_zone_numbers_local)

    if check_for_no_answer:
        main_window.destroy()
        return None

    # basic calculations and conversions
    sko_result_pos = sko_init(error_calculated_positions, peak_hkl)
    sko_result_intens = sko_init(error_calculated_intensities, peak_hkl)
    abs_err()
    message_for_frames()

    # removes the initial windows after selecting a file
    if main_window_destroy and not secondary_window_destroy:
        main_window.destroy()
    if secondary_window_destroy:
        secondary_window.destroy()

    secondary_window = Tk()
    secondary_window.title("Metrology")

    # disables the ability to zoom the page
    secondary_window.resizable(False, False)

    main_frame = LabelFrame(secondary_window)
    main_frame.pack(side=TOP)

    # creates the basic frames
    frame_for_checkbutton = LabelFrame(main_frame, bg='#dddddd')
    frame_for_down_buttons = LabelFrame(main_frame)

    # for standard deviation of intensities
    frame3 = LabelFrame(main_frame)
    # for standard deviation of positions
    frame4 = LabelFrame(main_frame)
    # for absolute error
    frame5 = LabelFrame(main_frame)

    frame_for_checkbutton.pack(side=RIGHT)
    frame_for_down_buttons.pack(side=BOTTOM)

    frame3.pack(side=TOP)
    frame4.pack(side=TOP)
    frame5.pack(side=TOP)

    # frames for standard deviation of intensities
    frame3_1 = LabelFrame(frame3)
    frame3_2 = LabelFrame(frame3)
    frame3_3 = LabelFrame(frame3)
    frame3_4 = LabelFrame(frame3)
    frame3_1.pack(side=TOP)
    frame3_2.pack(side=LEFT)
    frame3_3.pack(side=LEFT)
    frame3_4.pack(side=LEFT)

    # frames for standard deviation of positions
    frame4_1 = LabelFrame(frame4)
    frame4_2 = LabelFrame(frame4)
    frame4_3 = LabelFrame(frame4)
    frame4_4 = LabelFrame(frame4)
    frame4_1.pack(side=TOP)
    frame4_2.pack(side=LEFT)
    frame4_3.pack(side=LEFT)
    frame4_4.pack(side=LEFT)

    # frames for absolute error
    frame5_1 = LabelFrame(frame5)
    frame5_2 = LabelFrame(frame5)
    frame5_3 = LabelFrame(frame5)
    frame5_4 = LabelFrame(frame5)
    frame5_1.pack(side=TOP)
    frame5_2.pack(side=LEFT)
    frame5_3.pack(side=LEFT)
    frame5_4.pack(side=LEFT)

    hkl_for_frame = []

    def update_errors_for_frames():
        """Updates data in metrology results by adding variable triggers for checkboxes"""

        # adds a check for outputs over values; a variable for checkboxes, and the color of the checkbox
        for keys, values in sko_result_intens.items():
            if values != 'None information' and values['sko'] > 2:
                values['checkbutton_variable'] = False
                values['checkbutton_variable_for_def'] = IntVar()
            elif values == 'None information':
                pass
            else:
                values['checkbutton_variable'] = True
                values['checkbutton_variable_for_def'] = IntVar()

        for keys, values in sko_result_pos.items():
            if values != 'None information' and round(values['sko']) >= 0.02 and sko_result_pos[peak_hkl[keys]][
                        'metrology_var'] == 1:
                values['checkbutton_variable'] = False
                values['checkbutton_variable_for_def'] = IntVar()
            elif values == 'None information':
                pass
            else:
                values['checkbutton_variable'] = True
                values['checkbutton_variable_for_def'] = IntVar()

        timer_trig = 0
        for keys, values in absolute_error.items():
            if 'None information' not in values and values != [] and max(values['exp_data']) > 0.02:
                absolute_error[peak_hkl[timer_trig]]['checkbutton_variable'] = False
                absolute_error[peak_hkl[timer_trig]]['checkbutton_variable_for_def'] = IntVar()

            elif 'None information' in values or values == []:
                absolute_error[peak_hkl[timer_trig]] = 'None information'
            else:
                absolute_error[peak_hkl[timer_trig]]['checkbutton_variable'] = True
                absolute_error[peak_hkl[timer_trig]]['checkbutton_variable_for_def'] = IntVar()

            timer_trig += 1

    update_errors_for_frames()

    class Labels:
        """Frame configuration 'St.dev of intensities', 'St.dev of positions' and 'Absolute error'"""

        def __init__(self):
            # frame for standard deviation of intensities
            lbl_main_1 = Label(frame3_1, text='СКО Интенсивности:', width=63, bg='lightgreen')
            lbl_main_1.pack(side=TOP)

            lbl_min_3_1 = Label(frame3_2, text='минимальное:', width=20, bg='grey60')
            lbl_min_3_1.pack(side=TOP)
            self.lbl_real_min_3_1 = Label(frame3_2, width=20, bg='white')
            self.lbl_real_min_3_1.pack(side=TOP)
            self.lbl_hkl_min_3_1 = Label(frame3_2, width=20, bg='grey70')
            self.lbl_hkl_min_3_1.pack(side=TOP)

            lbl_avr_3_1 = Label(frame3_3, text='среднее:', width=20, bg='grey60')
            lbl_avr_3_1.pack(side=TOP)
            self.lbl_real_avr_3_1 = Label(frame3_3, width=20, bg='white')
            self.lbl_real_avr_3_1.pack(side=TOP)
            lbl_hkl_avr_3_1 = Label(frame3_3, text='<-- hkl -->', width=20, bg='grey70')
            lbl_hkl_avr_3_1.pack(side=TOP)

            lbl_max_3_1 = Label(frame3_4, text='максимальное:', width=20, bg='grey60')
            lbl_max_3_1.pack(side=TOP)
            self.lbl_real_max_3_1 = Label(frame3_4, width=20)
            self.lbl_real_max_3_1.pack(side=TOP)
            self.lbl_hkl_max_3_1 = Label(frame3_4, width=20, bg='grey70')
            self.lbl_hkl_max_3_1.pack(side=TOP)

            # frame for standard deviation of positions
            lbl_main_2 = Label(frame4_1, text='СКО Положения:', width=63, bg='lightgreen')
            lbl_main_2.pack(side=TOP)

            lbl_min_4_1 = Label(frame4_2, text='минимальное:', width=20, bg='grey60')
            lbl_min_4_1.pack(side=TOP)
            self.lbl_real_min_4_1 = Label(frame4_2, width=20, bg='white')
            self.lbl_real_min_4_1.pack(side=TOP)
            self.lbl_hkl_min_4_1 = Label(frame4_2, width=20, bg='grey70')
            self.lbl_hkl_min_4_1.pack(side=TOP)

            lbl_avr_4_1 = Label(frame4_3, text='среднее:', width=20, bg='grey60')
            lbl_avr_4_1.pack(side=TOP)
            self.lbl_real_avr_4_1 = Label(frame4_3, width=20, bg='white')
            self.lbl_real_avr_4_1.pack(side=TOP)
            lbl_hkl_avr_4_1 = Label(frame4_3, text='<-- hkl -->', width=20, bg='grey70')
            lbl_hkl_avr_4_1.pack(side=TOP)

            lbl_max_4_1 = Label(frame4_4, text='максимальное:', width=20, bg='grey60')
            lbl_max_4_1.pack(side=TOP)
            self.lbl_real_max_4_1 = Label(frame4_4, width=20)
            self.lbl_real_max_4_1.pack(side=TOP)
            self.lbl_hkl_max_4_1 = Label(frame4_4, width=20, bg='grey70')
            self.lbl_hkl_max_4_1.pack(side=TOP)

            # frame for absolute error
            lbl_main_3 = Label(frame5_1, text='Абсолютная погрешность:', width=63, bg='lightgreen')
            lbl_main_3.pack(side=TOP)

            lbl_min_5_1 = Label(frame5_2, text='минимальное:', width=20, bg='grey60')
            lbl_min_5_1.pack(side=TOP)
            self.lbl_real_min_5_1 = Label(frame5_2, width=20, bg='white')
            self.lbl_real_min_5_1.pack(side=TOP)
            self.lbl_hkl_min_5_1 = Label(frame5_2, width=20, bg='grey70')
            self.lbl_hkl_min_5_1.pack(side=TOP)

            lbl_avr_5_1 = Label(frame5_3, text='среднее:', width=20, bg='grey60')
            lbl_avr_5_1.pack(side=TOP)
            self.lbl_real_avr_5_1 = Label(frame5_3, width=20, bg='white')
            self.lbl_real_avr_5_1.pack(side=TOP)
            lbl_hkl_avr_5_1 = Label(frame5_3, text='<-- hkl -->', width=20, bg='grey70')
            lbl_hkl_avr_5_1.pack(side=TOP)

            lbl_max_5_1 = Label(frame5_4, text='максимальное:', width=20, bg='grey60')
            lbl_max_5_1.pack(side=TOP)
            self.lbl_real_max_5_1 = Label(frame5_4, width=20)
            self.lbl_real_max_5_1.pack(side=TOP)
            self.lbl_hkl_max_5_1 = Label(frame5_4, width=20, bg='grey70')
            self.lbl_hkl_max_5_1.pack(side=TOP)

        def labels_config(self):
            """Frame generation 'St.dev of intensities', 'St.dev of positions' and 'Absolute error'"""

            # variables to change the color and font of the maximum frame values when going out of bounds
            sko_int_color = 'green2'
            sko_pos_color = 'green2'
            sko_avr_color = 'green2'
            sko_int_color_txt = 'black'
            sko_pos_color_txt = 'black'
            sko_avr_color_txt = 'black'

            if sko_pos[2][1] == 0 and sko_int[2][1] == 0:
                abs_pos[0][1] = 0.0
                abs_pos[1] = 0.0
                abs_pos[2][1] = 0.0
                abs_pos[0][0] = 0.0
                abs_pos[2][0] = 0.0
                sko_pos[0][1] = 0.0
                sko_pos[1] = 0.0
                sko_pos[2][1] = 0.0
                sko_int[0][1] = 0.0
                sko_int[1] = 0.0
                sko_int[2][1] = 0.0

            else:
                if sko_int[2][1] > 2:
                    sko_int_color = 'red'
                    sko_int_color_txt = 'white'
                if sko_pos[2][1] > 0.02:
                    sko_pos_color = 'red'
                    sko_pos_color_txt = 'white'
                if abs_pos[2][1] > 0.02:
                    sko_avr_color = 'red'
                    sko_avr_color_txt = 'white'

            self.lbl_real_min_3_1.config(text=str(round(sko_int[0][1], 3)) + '%')
            self.lbl_hkl_min_3_1.config(text=str(sko_int[0][0]))
            self.lbl_real_avr_3_1.config(text=str(round(float(sko_int[1]), 3)) + '%')
            self.lbl_real_max_3_1.config(text=str(round(sko_int[2][1], 3)) + '%', bg=sko_int_color,
                                         fg=sko_int_color_txt)
            self.lbl_hkl_max_3_1.config(text=str(sko_int[2][0]))

            self.lbl_real_min_4_1.config(text=str(round(sko_pos[0][1], 4)))
            self.lbl_hkl_min_4_1.config(text=str(sko_pos[0][0]))
            self.lbl_real_avr_4_1.config(text=str(round(float(sko_pos[1]), 4)))
            self.lbl_real_max_4_1.config(text=str(round(sko_pos[2][1], 4)), bg=sko_pos_color, fg=sko_pos_color_txt)
            self.lbl_hkl_max_4_1.config(text=str(sko_pos[2][0]))

            self.lbl_real_min_5_1.config(text=str(round(abs_pos[0][1], 5)))
            self.lbl_hkl_min_5_1.config(text=str(abs_pos[0][0]))
            self.lbl_real_avr_5_1.config(text=str(round(float(abs_pos[1]), 5)))
            self.lbl_real_max_5_1.config(text=str(round(abs_pos[2][1], 5)), bg=sko_avr_color, fg=sko_avr_color_txt)
            self.lbl_hkl_max_5_1.config(text=str(abs_pos[2][0]))

    reload_x = Labels()
    reload_x.labels_config()

    def reload():
        for keys, values in absolute_error.items():
            if values != 'None information':
                # получает информацию по очереди о всех состояниях кнопок
                absolute_error[keys]['metrology_var'] = absolute_error[keys][
                    'checkbutton_variable_for_def'].get()
                sko_result_pos[keys]['metrology_var'] = absolute_error[keys][
                    'checkbutton_variable_for_def'].get()
                sko_result_intens[keys]['metrology_var'] = absolute_error[keys][
                    'checkbutton_variable_for_def'].get()
        message_for_frames()
        reload_x.labels_config()

    # Check positions and converts to a single form
    from main_data_processing import nist_peak_position

    timer = 0
    max_len_of_key = []

    for key, value in sko_result_pos.items():
        max_len_of_key.append(len(key))
    max_len_of_key = max(max_len_of_key)

    for key, value in sko_result_pos.items():
        if len(key) < max_len_of_key:
            hkl_for_frame.append(' ' + key + ' ' * (int(max_len_of_key) - len(key) + 2))
        else:
            hkl_for_frame.append(' ' + key + ' ')

        # if there is no information, then the frame is deactivated, if over 100, then a little less gaps
        if len(value) <= 15:

            if float(value['exp_data'][1]) >= 100:
                if absolute_error[key]['checkbutton_variable'] and sko_result_pos[key][
                            'checkbutton_variable'] and sko_result_intens[key]['checkbutton_variable']:
                    chckbtn = Checkbutton(frame_for_checkbutton,
                                          text=hkl_for_frame[-1] + ' -   ' + nist_peak_position[timer] + '°',
                                          bg='#dddddd',
                                          variable=absolute_error[peak_hkl[timer]][
                                              'checkbutton_variable_for_def'],
                                          command=reload)
                    chckbtn.grid(row=timer, sticky=W, pady=0)
                    Checkbutton.select(chckbtn)
                else:
                    chckbtn = Checkbutton(frame_for_checkbutton,
                                          text=hkl_for_frame[-1] + ' -   ' + nist_peak_position[timer] + '°',
                                          fg='red', bg='#dddddd', command=reload,
                                          variable=absolute_error[peak_hkl[timer]][
                                              'checkbutton_variable_for_def'])
                    chckbtn.grid(row=timer, sticky=W, pady=0)
                    Checkbutton.select(chckbtn)
            else:
                if absolute_error[key]['checkbutton_variable'] and sko_result_pos[key][
                            'checkbutton_variable'] and sko_result_intens[key]['checkbutton_variable']:
                    chckbtn = Checkbutton(frame_for_checkbutton,
                                          text=hkl_for_frame[-1] + ' -   ' + nist_peak_position[timer] + '°  ',
                                          bg='#dddddd',
                                          command=reload,
                                          variable=absolute_error[peak_hkl[timer]][
                                              'checkbutton_variable_for_def'])
                    chckbtn.grid(row=timer, sticky=W, pady=0)
                    Checkbutton.select(chckbtn)
                else:
                    chckbtn = Checkbutton(frame_for_checkbutton,
                                          text=hkl_for_frame[-1] + ' -   ' + nist_peak_position[timer] + '°  ',
                                          fg='red', bg='#dddddd', command=reload, variable=absolute_error
                                          [peak_hkl[timer]]['checkbutton_variable_for_def'])
                    chckbtn.grid(row=timer, sticky=W, pady=0)
                    Checkbutton.select(chckbtn)
        else:
            if float(nist_peak_position[timer]) >= 100:
                chckbtn = Checkbutton(frame_for_checkbutton,
                                      text=hkl_for_frame[-1] + ' -   ' + nist_peak_position[timer] + '°',
                                      bg='grey70', state=DISABLED)
                chckbtn.grid(row=timer, sticky=W, pady=0)
            else:
                chckbtn = Checkbutton(frame_for_checkbutton,
                                      text=hkl_for_frame[-1] + ' -   ' + nist_peak_position[timer] + '°  ', bg='grey70',
                                      state=DISABLED)
                chckbtn.grid(row=timer, sticky=W, pady=0)

        timer += 1

    class Info:
        def __init__(self):
            self.open_close_trigger = 0
            from main_data_processing import directory_expand
            from main_data_processing import srm_name
            from main_data_processing import number_of_diffractograms
            global secondary_window
            self.information_frame = LabelFrame(secondary_window)

            self.main_info_lbl = Label(self.information_frame, width=64, justify=LEFT, wraplength=435, anchor=NW,
                                       relief=RIDGE, text=' Путь к файлу:\n  - ' + str(directory_expand) +
                                       '\n\n Стандарт образца:\n  - NIST ' + str(srm_name) +
                                       '\n\n Количество снятых дифрактограмм:\n  - ' + str(number_of_diffractograms) +
                                       '\n\n Версия формул:\n  - март 2022', pady=8)

            self.main_info_lbl.pack(side=LEFT, anchor=NW, fill=Y)
            self.info_for_checkbuttons = Label(self.information_frame, width=20, bg='#dddddd')
            self.info_for_checkbuttons.pack(side=LEFT, fill=Y)
            self.black_chckbtn = Label(self.info_for_checkbuttons, bg='#dddddd', justify=CENTER, text='ЧЕРНЫЙ:',
                                       fg='black', relief=GROOVE)
            self.black_chckbtn.pack(side=TOP, fill=X)
            self.black_chckbtn_info = Label(self.info_for_checkbuttons, bg='#dddddd', justify=CENTER,
                                            text='активное поле в \nрамках нормального\nотклонения',
                                            fg='black', relief=GROOVE, borderwidth=1, pady=3)
            self.black_chckbtn_info.pack(side=TOP, fill=X)
            self.gray_chckbtn = Label(self.info_for_checkbuttons, bg='#dddddd', justify=CENTER, text='СЕРЫЙ:',
                                      fg='gray30', relief=GROOVE)
            self.gray_chckbtn.pack(side=TOP, fill=X)
            self.gray_chckbtn_info = Label(self.info_for_checkbuttons, bg='#dddddd', justify=CENTER,
                                           text='отсутствует\nинформация о пике', relief=GROOVE, borderwidth=1, pady=3)
            self.gray_chckbtn_info.pack(side=TOP, fill=X)
            self.red_chckbtn_info = Label(self.info_for_checkbuttons, bg='#dddddd', justify=CENTER,
                                          text='КРАСНЫЙ:', fg='red', width=18, relief=GROOVE)
            self.red_chckbtn_info.pack(side=TOP, fill=X)

            self.red_chckbtn_info = Label(self.info_for_checkbuttons, bg='#dddddd', justify=CENTER, width=18,
                                          relief=GROOVE, text='выход за max\nотклонение', borderwidth=1, pady=3)
            self.red_chckbtn_info.pack(side=TOP, fill=Y)

        def reload(self):
            global open_close_trigger
            self.open_close_trigger += 1
            if self.open_close_trigger == 1:
                self.information_frame.pack(side=TOP)

            else:
                self.open_close_trigger = 0
                self.information_frame.pack_forget()

    info_reload = Info()

    def chose_another_file():
        global secondary_window_destroy
        secondary_window_destroy = True
        open_file()

    # new interface buttons
    reload_btn = Button(frame_for_down_buttons, text="Создать отчёт", width=20)
    again_btn = Button(frame_for_down_buttons, text="Выбрать другой файл", width=20, command=chose_another_file)
    info_btn = Button(frame_for_down_buttons, text="Информация", width=20, command=info_reload.reload)
    again_btn.pack(side=LEFT)
    info_btn.pack(side=LEFT)
    reload_btn.pack(side=BOTTOM)

    # sets the size of the window and places it in the center of the screen
    secondary_window.update_idletasks()
    secondary_w = secondary_window.geometry()
    secondary_w = secondary_w.split('+')
    secondary_w = secondary_w[0].split('x')
    width_secondary_window = int(secondary_w[0])
    height_secondary_window = int(secondary_w[1])

    secondary_width = secondary_window.winfo_screenwidth()
    secondary_height = secondary_window.winfo_screenheight()
    secondary_width = secondary_width // 2
    secondary_height = secondary_height // 2
    secondary_width = secondary_width - width_secondary_window // 2
    secondary_height = secondary_height - height_secondary_window // 2
    secondary_window.geometry('+{}+{}'.format(secondary_width, secondary_height))


main_window = Tk()
main_window.title("Metrology")

# disables the ability to zoom the page
main_window.resizable(False, False)

# frame for the main interface
frame1 = LabelFrame(main_window)
frame1.pack(side=LEFT)

# outputs the information about the absolute error in the GUI
start_btn = Button(frame1, text="Выбрать \nи обработать файл  ", relief=GROOVE, width=18, command=open_file)
start_btn.pack(side=TOP)
search_area_button = Button(frame1, text="Область поиска\n данных в документе", relief=GROOVE, width=18,
                            command=change_search_area)
search_area_button.pack(side=TOP)

# sets the size of the window and places it in the center of the screen
main_window.update_idletasks()  # Updates information after all frames are created
s = main_window.geometry()
s = s.split('+')
s = s[0].split('x')
width_main_window = int(s[0])
height_main_window = int(s[1])

w = main_window.winfo_screenwidth()
h = main_window.winfo_screenheight()
w = w // 2
h = h // 2
w = w - width_main_window // 2
h = h - height_main_window // 2
main_window.geometry('+{}+{}'.format(w, h))

main_window.mainloop()
