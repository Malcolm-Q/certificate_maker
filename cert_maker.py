from pypdf import PdfReader, PdfWriter
from pandas import read_csv
import tkinter as tk
from tkinter import filedialog
from subprocess import Popen
from os import getcwd

# ugly global var definition
csv_disclaimer, open_csv_button, csv_label, done_button_date, date_label, date_entry, date, file_suffix, reader, name_list, pdf_label, open_pdf_button, pdf_disclaimer, done_button = None,None,None,None,None,None,None,None,None,None,None,None,None,None

# set up window
window = tk.Tk()
window.title("Certificate Maker")
window.geometry("400x250")

# set up first screen. Once a value is submitted a new screen will appear to reduce clutter.
suffix_label = tk.Label(window, text="\nEnter file suffix:\n\n\n")
suffix_label.pack()

suffix_entry = tk.Entry(window,width=20)
suffix_entry.pack()

# The first screen asks for the file suffix
def main():
    global done_button
    done_button = tk.Button(window, text=' CONFIRM SUFFIX ', command=set_suffix)
    done_button.configure(bg='#6cd929',fg='black')
    done_button.pack()
    window.mainloop()

# save suffix to global var, setup date screen.
def set_suffix():
    global file_suffix, date_entry, date_label, done_button_date

    file_suffix = suffix_entry.get()

    # destroy suffix fields
    suffix_label.destroy()
    suffix_entry.destroy()
    done_button.destroy()

    # create date input fields
    date_label = tk.Label(window, text="\nEnter graduation date:\n\n\n")
    date_label.pack()

    date_entry = tk.Entry(window,width=20)
    date_entry.pack()

    done_button_date = tk.Button(window, text=' CONFIRM DATE ', command=set_date)
    done_button_date.configure(bg='#6cd929',fg='black')
    done_button_date.pack()

# save date, setup csv selection
def set_date():
    global date, date_entry, date_label, done_button_date, csv_label, open_csv_button, csv_disclaimer

    # destroy date fields
    date = date_entry.get()
    date_entry.destroy()
    date_label.destroy()
    done_button_date.destroy()

    # setup csv fields
    csv_label = tk.Label(window, text='\nSelect CSV containing names\nMust have "NAMES" column!\n\n\n')
    csv_label.pack()

    open_csv_button = tk.Button(window, text="Open .csv", command=open_csv)
    open_csv_button.configure(bg='#6cd929',fg='black')
    open_csv_button.pack()

    csv_disclaimer = tk.Label(window, text='')
    csv_disclaimer.pack()

# try load csv, setup pdf selection
def open_csv():
    global name_list, csv_label, csv_disclaimer, open_csv_button, pdf_disclaimer, pdf_label, open_pdf_button

    # only allow user to open .csv files
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    try:
        name_list = read_csv(file_path)
        name_list = name_list['NAMES'].tolist()
        csv_disclaimer.config(text=f"Loaded File: {file_path}")
    except:
        # if there's an error inform user and return so they can try again.
        csv_disclaimer.config(text="Failed to load file.")
        return
    
    # destroy fields
    csv_disclaimer.destroy()
    csv_label.destroy()
    open_csv_button.destroy()

    # setup pdf selection
    pdf_label = tk.Label(window, text='\nSelect .pdf template\nMust have "Name" and "Granted on" field!\n\n\n')
    pdf_label.pack()

    open_pdf_button = tk.Button(window, text="Open .pdf", command=open_pdf)
    open_pdf_button.configure(bg='#6cd929',fg='black')
    open_pdf_button.pack()

    pdf_disclaimer = tk.Label(window, text='')
    pdf_disclaimer.pack()

# try load pdf, setup final button
def open_pdf():
    # only allow user to open .pdf files
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    global reader, pdf_disclaimer, pdf_label, open_pdf_button
    try:
        reader = PdfReader(file_path)
        pdf_disclaimer.config(text=f'Loaded File: {file_path}')
    except:
        # return if error to allow user to retry
        pdf_disclaimer.config(text='Failed to load file.')
        return
    
    # destroy fields
    pdf_disclaimer.destroy()
    pdf_label.destroy()
    open_pdf_button.destroy()

    # setup final button that writes certs
    blank = tk.Label(window, text='\n\n\n')
    blank.pack()

    write_cert_button = tk.Button(window, text="\n   Write PDFs   \n", command=write_certs)
    write_cert_button.configure(bg='#6cd929',fg='black')
    write_cert_button.pack()

# write certificates using names from csv, user entered date, and pdf template. Save to working dir and open working dir.
def write_certs():
    global reader, name_list, date

    # init pypdf
    writer = PdfWriter()
    date = date
    page = reader.pages[0]
    writer.add_page(page)

    for name in name_list:
        # create cert for each name and save
        writer.update_page_form_field_values(
            writer.pages[0], {"Name": name,'Granted on':date}
        )

        with open(f'{name}_{file_suffix}.pdf', 'wb') as output_stream:
            writer.write(output_stream)
    # open file explorer and kill program.
    Popen(f'''explorer /select,"{getcwd()}\{name_list[0]}_{file_suffix}.pdf"''')
    window.destroy()



if __name__ == '__main__':
    main()