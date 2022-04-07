import os
import tkinter
from nrfreq import generate_nrfreqs
from ext_qci import generate_extqci
from cosite_4g_4g import generate_4g_4g_cosite
from cosite_5g_5g import generate_5g_5g_cosite
from icon import Icon
from os.path import exists
import threading


proc_dict = {generate_nrfreqs: "Generating NRFREQ", generate_extqci: "Generating EXTQCI",
             generate_4g_4g_cosite: "Generating Cosites 4G-4G", generate_5g_5g_cosite: "Generating Cosites 5G-5G"}


def validate_input(tbinput):
    try:
        inputs = tbinput.split(";")
        for iinput in inputs:
            if len(iinput) != 8:
                return False
        return True
    except:
        return False


def global_process():
    threading.Thread(target=mo_process).start()


def mo_process():
    lbl_result.grid_forget()
    tbinput = me_id_tb.get().strip()
    result = validate_input(tbinput)

    if not result:
        result_text.set("Input error")
        lbl_result.grid(row=4, column=0, sticky=tkinter.W)
        return
    if not exists('data/input.xlsx'):
        result_text.set("Cannot find input.xlsx in data folder")
        lbl_result.grid(row=4, column=0, sticky=tkinter.W)
        return
    process["state"] = "disabled"
    lbl_result.grid(row=4, column=0, sticky=tkinter.W)
    if allp.get():
        try:
            result_text.set("Generating NRFREQ")
            workbook = generate_nrfreqs(tbinput)
            result_text.set("Generating EXTQCI")
            workbook = generate_extqci(tbinput, workbook)
            result_text.set("Generating Cosites 4G-4G")
            workbook = generate_4g_4g_cosite(tbinput, workbook)
            result_text.set("Generating Cosites 5G-5G")
            generate_5g_5g_cosite(tbinput, workbook, True)
        except Exception as e:
            result_text.set(e)
            process["state"] = "normal"
            return
    else:
        steps = [generate_nrfreqs, generate_extqci, generate_4g_4g_cosite, generate_5g_5g_cosite]
        selected = [nrfreq.get(), ext_qci.get(), cos_4g_4g.get(), cos_5g_5g.get()]
        process_steps = [step for idx, step in enumerate(steps) if selected[idx]]
        if len(process_steps) == 0:
            result_text.set("Please, select at least 1 process")
            process["state"] = "normal"
            return

        workbook = None
        try:
            for step in process_steps:
                if len(process_steps) == 1:
                    result_text.set(proc_dict[step])
                    step(tbinput, None, True)
                elif step == process_steps[-1]:
                    result_text.set(proc_dict[step])
                    step(tbinput, workbook, 1)
                elif step == process_steps[0]:
                    result_text.set(proc_dict[step])
                    workbook = step(tbinput)
                else:
                    result_text.set(proc_dict[step])
                    workbook = step(tbinput, workbook, 0)
        except Exception as e:
            result_text.set(e)
            process["state"] = "normal"
            return

    result_text.set("Process finished")
    process["state"] = "normal"


def indiv_selected():
    allp.set(value=0)


def select_all():
    nrfreq.set(value=0)
    ext_qci.set(value=0)
    cos_4g_4g.set(value=0)
    cos_5g_5g.set(value=0)


root = tkinter.Tk()
root.geometry("285x200")
root.resizable(False, False)
root.title("Baseline Automation Tool")

result_text = tkinter.StringVar()
nrfreq = tkinter.IntVar()
ext_qci = tkinter.IntVar()
cos_4g_4g = tkinter.IntVar()
cos_5g_5g = tkinter.IntVar()
allp = tkinter.IntVar(value=1)
lbl = tkinter.Label(root, text="ME ID list", pady=20, padx=20)
lbl.grid(row=0, column=0, sticky=tkinter.W)
me_id_tb = tkinter.Entry(root, borderwidth=2)
me_id_tb.grid(row=0, column=1, sticky=tkinter.W)
# me_id_tb.insert(0, "PVAX0352;PVAX0010")
# me_id_tb.insert(0, "Enter ME ID")
chb_nrfreq = tkinter.Checkbutton(root, text="NRFREQ", variable=nrfreq, padx=20, command=indiv_selected)
chb_nrfreq.grid(row=1, column=0, sticky=tkinter.W)
chb_ext_qci = tkinter.Checkbutton(root, text="EXTQCI", variable=ext_qci, command=indiv_selected)
chb_ext_qci.grid(row=1, column=1, sticky=tkinter.W)
chb_cos_4g_4g = tkinter.Checkbutton(root, text="Cosites 4G-4G", variable=cos_4g_4g, padx=20, command=indiv_selected)
chb_cos_4g_4g.grid(row=2, column=0, sticky=tkinter.W)
chb_cos_5g_5g = tkinter.Checkbutton(root, text="Cosites 5G-5G", variable=cos_5g_5g, command=indiv_selected)
chb_cos_5g_5g.grid(row=2, column=1, sticky=tkinter.W)
chb_all = tkinter.Checkbutton(root, text="All", variable=allp, padx=20, command=select_all)
chb_all.grid(row=3, column=0, sticky=tkinter.W)
lbl_result = tkinter.Label(root, text="", padx=20, textvariable=result_text, wraplength=100)
lbl_result.grid(row=4, column=0, sticky=tkinter.W)
process = tkinter.Button(root, text="Process", command=global_process)
process.grid(row=4, column=1, sticky=tkinter.W)

lbl_result.grid_forget()
Icon()
root.iconbitmap('logo.ico')
os.remove("logo.ico")
root.mainloop()

# PVAX0843;PVAX0774;PVAX0302;PVAX0844