from tkinter import *
from tkinter.ttk import *
from string import Template
from tkinter import ttk
from tkinter import messagebox
import os
import shutil
from timeit import default_timer as timer
from PIL import Image
from tkinter import filedialog

# tkinter size and title section
window = Tk()
window.title("Toolbox v1.01")
window.geometry("325x235")

# tkinter tabs section
tab_control = ttk.Notebook(window)
generate_cps_tab = ttk.Frame(tab_control)
employee_code_tab = ttk.Frame(tab_control)
multiple_labels = ttk.Frame(tab_control)
nxview = ttk.Frame(tab_control)
tab_control.add(generate_cps_tab, text="Generator CPS")
tab_control.add(employee_code_tab, text="Kody pracowników")
tab_control.add(multiple_labels, text="Wiele etykiet")
tab_control.add(nxview, text="NxView")

# ###################################################################################################################
# Code for generate_cps_tab
# ###################################################################################################################

project_list = Combobox(generate_cps_tab)
project_list["values"] = ("Company", "Company2", "Company3")
project_list.grid(column=0, row=2, padx=5, pady=5)

project_label = Label(generate_cps_tab, text="Wybór projektu:")
project_label.grid(column=0, row=1, padx=5, pady=5)

harness_id_label = Label(generate_cps_tab, text="Numer wiązki:")
harness_id_label.grid(column=1, row=1, padx=5, pady=5)
harness_id_entry = Entry(generate_cps_tab, width=22)
harness_id_entry.grid(column=1, row=2, padx=5, pady=5)


def len_error(entry_num, char_num):
    messagebox.showerror(title="Wystąpił błąd", message=f"Długość {entry_num} pola przekracza {char_num} znaków")

# function used by button - all commands for generate a .prn label
def generate():
    text1 = first_line_entry.get()
    if len(text1) >= 9:
        len_error("pierwszego", "8")
        return
    text2 = second_line_entry.get()
    if len(text2) >= 17:
        len_error("drugiego", "16")
        return
    text3 = third_line_entry.get()
    if len(text3) >= 17:
        len_error("trzeciego", "16")
        return

    project_name = project_list.get()
    try:
        if text3 != "":
            if text2 == "":
                messagebox.showerror(title="Wystąpił błąd", message="Nie zostały wypełnione wszystkie pola")
                return
            if text1 == "":
                messagebox.showerror(title="Wystąpił błąd", message="Nie zostały wypełnione wszystkie pola")
                return

            example = project_name + "3.prn"
        elif text2 != "":
            if text1 == "":
                messagebox.showerror(title="Wystąpił błąd", message="Nie zostało wypełnione pole nr 1.")
                return

            example = project_name + "2.prn"
        elif text1 != "":
            example = project_name + "1.prn"

        harness_id = str(harness_id_entry.get())
        path_to_folder = "c:\\Users\\user\\Desktop"
        path = os.path.join(path_to_folder, harness_id)
        if harness_id != "":
            try:
                os.makedirs(path)
            except FileExistsError:
                pass
        with open(example) as f:
            input_file = f.read()
        with open(text1 + ".prn", "w", encoding="utf-8") as wf:
            wf.write(Template(input_file).safe_substitute(first=text1, second=text2, third=text3))
        first_line_entry.delete(0, END)
        second_line_entry.delete(0, END)
        third_line_entry.delete(0, END)
    except UnboundLocalError:
        messagebox.showerror(title="Wystąpił błąd", message="Etykieta nie może być pusta")
        return
    except FileNotFoundError:
        messagebox.showerror(title="Wystąpił błąd", message="Nie wybrano projektu.")
        return
    final_name = text1 + ".prn"
    destination = "c:\\Users\\user\\Desktop" + "\\" + harness_id + "\\"
    program_dir = os.getcwd()
    if os.path.isfile(destination + "\\" + final_name):
        replace_result = messagebox.askyesno(title="Wystąpił błąd", message="CPS o danym numerze istnieje. Podmienić?")
        if replace_result is True:
            os.remove(destination + final_name)
            shutil.move(program_dir + "\\" + final_name, destination)
            return
        else:
            return
    shutil.move(program_dir + "\\" + final_name, destination)


first_line = Label(generate_cps_tab, text="Piewsza linia CPS:")
first_line.grid(column=0, row=4)
first_line_entry = Entry(generate_cps_tab, width=12)
first_line_entry.grid(column=1, row=4, padx=5, pady=5)

second_line = Label(generate_cps_tab, text="Druga linia CPS:")
second_line.grid(column=0, row=5)
second_line_entry = Entry(generate_cps_tab, width=22)
second_line_entry.grid(column=1, row=5, padx=5, pady=5)

third_line = Label(generate_cps_tab, text="Trzecia linia CPS:")
third_line.grid(column=0, row=6)
third_line_entry = Entry(generate_cps_tab, width=22)
third_line_entry.grid(column=1, row=6, padx=5, pady=5)

generate_button = Button(generate_cps_tab, text="Generuj CPS", command=generate)
generate_button.grid(column=1, row=7, pady=15)

# ###################################################################################################################
# Code for operator labels
# ###################################################################################################################


def clean():
    employee_code_entry.delete(0, END)
    employee_name_entry.delete(0, END)

#function used by generate button - for generate a employee label.
def employee_generate():
    name = employee_name_entry.get()
    code = employee_code_entry.get()
    if code == "":
        messagebox.showerror(title="Wystąpił błąd", message="Kod pracownika nie może być pusty")
        return
    if name == "":
        messagebox.showerror(title="Wystąpił błąd", message="Imię i nazwisko pracownika nie może być puste.")
        return

    try:
        type(code) != int(employee_code_entry.get())
    except ValueError:
        clean()
        messagebox.showerror(title="Wystąpił błąd", message="Kod pracownika powinien składać się z samych liczb.")
        return

    if not name.replace(" ", "").isalpha():
        messagebox.showerror(title="Wystąpił błąd", message="Imię i nazwisko pracownika zawiera nieprawidłowe znaki.")
        return

    with open("employee_code.prn") as f:
        input_file = f.read()
    with open(str(code) + ".prn", "w", encoding="utf-8") as wf:
        wf.write(Template(input_file).safe_substitute(first=code, second=name))

    path_to_folder = "c:\\Desktop"
    path = os.path.join(path_to_folder, "Kody Pracowników")
    try:
        os.makedirs(path)
    except FileExistsError:
        pass
    try:
        shutil.move(os.getcwd() + "\\" + code+".prn", path + "\\")
        clean()
    except shutil.Error:
        ask_employee_exist = messagebox.askyesno(title="Wystąpił błąd", message="CPS o danym numerze "
                                                                                "istnieje. Podmienić?")
        if ask_employee_exist is True:
            os.remove(path + "\\" + code + ".prn")
            shutil.move(os.getcwd() + "\\" + code+".prn", path + "\\")
            clean()
            return
        else:
            return


employee_code = Label(employee_code_tab, text="Kod pracownika:")
employee_code.grid(column=0, row=0, padx=5, pady=30)
employee_code_entry = Entry(employee_code_tab, width=6, justify="center")
employee_code_entry.grid(column=1, row=0, padx=5, pady=10)

employee_name = Message(employee_code_tab, text="Imię i nazwisko pracownika:", width=90)
employee_name.grid(column=0, row=2, padx=5, pady=10)
employee_name_entry = Entry(employee_code_tab, width=22, justify="center")
employee_name_entry.grid(column=1, row=2, padx=5, pady=10)

generate_button = Button(employee_code_tab, text="Generuj etykiete", command=employee_generate)
generate_button.grid(column=1, row=8)

# ###################################################################################################################
# code for multiple labels
# ###################################################################################################################

# code used by generate button - for generate a specific number of labels - based on user input.
def multiple_generate():
    try:
        if begin_entry.get() == "":
            messagebox.showerror(title="Wystąpił błąd", message="Drugie pole nie może być puste.")
            return
        quantity = int(quantity_entry.get()) + 1
        begin1 = begin_entry.get()
        project_name = project_list_labels.get()
        quantity_counter = 0
        example1 = project_name + "_multiple.prn"
        path_to_folder = "c:\\Desktop"
        path = os.path.join(path_to_folder, "Wygenerowane CPS - " + str(begin1))
        start_timer = timer()
        while quantity != int(quantity_counter):
            final_name = begin1 + str(quantity_counter) + ".prn"
            with open(example1) as f:
                input_file = f.read()
            with open(begin1 + str(quantity_counter) + ".prn", "w", encoding="utf-8") as wf:
                wf.write(Template(input_file).safe_substitute(first=begin1 + str(quantity_counter)))

            try:
                os.makedirs(path)
            except FileExistsError:
                pass
            try:
                shutil.move(os.getcwd() + "\\" + final_name, path + "\\")
            except shutil.Error:
                messagebox.showerror(title="Wystąpił błąd", message="Folder z wygenerowanymi CPS'ami istnieje.")
                return

            quantity_counter += 1
    except FileNotFoundError:
        messagebox.showerror(title="Wystąpił błąd", message="Nie wybrano projektu.")
        return

    except ValueError:
        messagebox.showerror(title="Wystąpił błąd", message="Nieprawidłowa ilość sztuk.")
        return
    stop_timer = timer()
    messagebox.showinfo(title="Gotowe", message=f"Wygenerowano {quantity_counter -1} etykiet w"
                                                f" {round(stop_timer - start_timer, 3)} sekundy ")


project_list_labels = Combobox(multiple_labels)
project_list_labels["values"] = ("Bobcat", "Projekt2")
project_list_labels.grid(column=0, row=2, padx=5, pady=5)
project_label = Label(multiple_labels, text="Wybór projektu:")
project_label.grid(column=0, row=1, padx=5, pady=5)

quantity_label = Label(multiple_labels, text="Ilość sztuk: ")
quantity_label.grid(column=0, row=3, pady=10)
quantity_entry = Entry(multiple_labels, width=12)
quantity_entry.grid(column=1, row=3)

begin = Label(multiple_labels, text="Początek np ''CRP'': ")
begin.grid(column=0, row=4, pady=5)
begin_entry = Entry(multiple_labels, width=12)
begin_entry.grid(column=1, row=4)

generate_button = Button(multiple_labels, text="Generuj CPS", command=multiple_generate)
generate_button.grid(column=1, row=8)

# ###################################################################################################################
# Code for NxView tab
# ###################################################################################################################


def imageresize():
    images_dir = filedialog.askdirectory()
    counter = 0
    files = os.listdir(images_dir)
    start_timer = timer()
    for i in files:
        myimg = Image.open(images_dir + "/" + i)
        values = myimg.size
        width = int(values[0] * 0.22)
        height = int(values[1] * 0.22)
        newimg = myimg.resize((width, height))
        name = str(counter) + ".jpg"
        newimg.save(name)
        path_to_save = "C:\\Users\\Filip Laszczak\\Desktop\\"
        path2 = os.path.join(path_to_save, "Zdjęcia NxView")
        try:
            os.mkdir(path2)
        except FileExistsError:
            pass
        try:
            shutil.move(os.getcwd() + "//" + name, "C:\\Users\\Filip Laszczak\\Desktop\\Zdjęcia NxView")
        except shutil.Error:
            messagebox.showerror(title="Wystąpił błąd", message="Folder docelowy (Zdjęcia NxView na pulpicie)"
                                                                " nie jest pusty.")
            return
        counter += 1
    stop_timer = timer()
    messagebox.showinfo(title="Gotowe", message=f"Wygenerowano {counter} zdjęć w"
                                                f" {round(stop_timer - start_timer, 3)} sekundy ")


select_folder_label = Label(nxview, text="Zmniejszanie zdjęć:")
select_folder_label.grid(column=0, row=3, padx=5, pady=15)
select_folder = Button(nxview, text="Wybierz folder", command=imageresize)
select_folder.grid(column=0, row=4, padx=5)

# tkinter control section
tab_control.grid(column=0, row=0)
window.mainloop()
