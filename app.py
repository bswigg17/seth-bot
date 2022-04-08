import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import * 
import sys, os
from modules.sethbot import seth_bot
import time

# from modules.seth_bot import SethBot


class App(tk.Tk):

    def __init__(self):
        self.__components = { 'Sprint Status Excel': 'Upload File',
                                    'Bug Export Excel': 'Upload File',
                                    'Story Export Excel': 'Upload File',
                        }
        self.sprint_week = None
        self.new_sprint = None
        self.__files = dict()
        super().__init__()
        self.config_root()

    def config_root(self):
        self.resizable(width=False, height=False)
        self.config_title()
        self.config_size()
        self.config_header()
        self.pack_components()
        self.button_frame()

    def config_title(self):
        self.title('Seth Bot')

    def config_size(self):
        self.geometry('600x520')

    def config_header(self):
        ttk.Label(self, text="Upload Files", font=("Arial", 30)).pack(pady=25)

    def pack_components(self):
        for label, button in self.__components.items():
            self.pack_componenet(label, button)
        self.pack_week()
        self.pack_new_sprint()

    def pack_componenet(self, label, button):
        self.pack_label(label)
        self.pack_button(label, button)

    def pack_week(self):
        ttk.Label(self, text='Sprint Week', font=('Arial', 16)).pack(pady=(25, 0))
        self.sprint_week = ttk.Entry(self)
        self.sprint_week.pack(pady=(10, 0))

    def pack_new_sprint(self):
        self.new_sprint = tk.IntVar()
        ttk.Checkbutton(self, text='New Sprint', variable=self.new_sprint, onvalue=1, offvalue=0).pack(pady=(10, 0))

    def pack_label(self, label):
        ttk.Label(self, text=label, font=("Arial", 16)).pack(pady=(25, 0))

    def pack_button(self, label, button):
        self.__components[label] = ttk.Button(self, text=button, command=lambda: self.file_upload(label))
        self.__components[label].pack(pady=0)
    
    def file_upload(self, label):
        try:
            self.__files[label] = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx")])
            if self.__files[label] == '':
                pass 
            else:
                self.__components[label].config(state='disabled')
                self.__components[label]['text'] = 'Uploaded'
                self.set_options_label_text('Options...')
        except Exception as E:
            self.set_options_label_text(E)

    def button_frame(self):
        self.__options_label = ttk.Label(self, text="Options...", font=("Arial", 10))
        self.__options_label.pack(pady=(50, 0))
        ttk.Button(self, text='Clear', command=lambda: self.file_clear()).pack(side=LEFT, padx=(202, 0))
        ttk.Button(self, text='Submit', command=lambda: self.file_submit()).pack(side=RIGHT, padx=(0, 202))
        
        
    def file_clear(self):
        self.set_options_label_text('Options...')
        if len(self.__components) > 0:
            for button in self.__components.values():
                button.config(state='enabled')
                button['text'] = 'Upload File'
            self.sprint_week.delete(0, tk.END)
            self.__files = dict()

    def file_submit(self):
        if len(self.__files) == 3 and (self.sprint_week.get() != ""):
            self.set_options_label_text('Generating Report...')
            #Seth Bot Object Here (sprint, backlog, bugs, week)
            results = seth_bot(*[file for file in self.__files.values()], week=int(self.sprint_week.get()), new_sprint=self.new_sprint.get())
            if results:
                self.set_options_label_text('DONE.')
                time.sleep(3)
                self.file_clear()
            else: 
                self.set_options_label_text('Error.')
        else:
            self.set_options_label_text("Error: Did you upload 3 files?")

    def set_options_label_text(self, text):
        self.__options_label['text'] = text


        

if __name__ == "__main__":
    app = App()
    app.mainloop()