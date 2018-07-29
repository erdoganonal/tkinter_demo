import time
from tkinter import *
from tkinter import filedialog, messagebox, ttk
import threading

EXCEL_FILE_TYPES = [("Excel files", "*.xlsx *.xls")]

DEFAULT_GET_FILE_TEXT = "Please select excel file to convert"
DEFAULT_GET_DIRECTORY_TEXT = "Please select directory to export file"
TITLE = "ISD 2-3 OVERTIME EXCEL GENERATOR"

WINDOW_WIDTH = 350
WINDOW_HEIGHT = 200
PROGRESS_BAR_WIDTH = 100
BUTTON_WIDTH = 30
BUTTON_HEIGHT = 30
BUTTON_LABEL_SPACE = 5
PAD_X = 10
PAD_Y = 5
X_POSITION = 10
Y_POSITION = 10
NEW_LINE_HEIGHT = 50

LABEL_SETTINGS = {
    'anchor': W,
    'font': 10,
    'relief': RIDGE,
    'wraplength': 0,
    'height': 1,
    'padx': PAD_X,
    'pady': PAD_Y,
    'bg': '#80ced6',
}


class ExcelConverterGui(object):
    def __init__(self):
        self.main_window = self.create_main_window()
        self.get_file_label_text = StringVar()
        self.get_directory_label_text = StringVar()
    
        self.group_count = 0
        self.add_items()

    def start_gui(self):
        self.main_window.mainloop()

    @staticmethod
    def create_main_window():
        main_window = Tk()
        main_window.title(TITLE)
        main_window.geometry("{width}x{height}".format(width=WINDOW_WIDTH, height=WINDOW_HEIGHT))

        return main_window

    @staticmethod
    def _warning(title, message):
        messagebox.showinfo(title, message)

    @staticmethod
    def get_file_path(label_text):
        filename = filedialog.askopenfilename(filetypes=EXCEL_FILE_TYPES)
        if filename:
            label_text.set(filename)

    @staticmethod
    def get_export_path(label_text):
        export_path = filedialog.askdirectory()
        if export_path:
            label_text.set(export_path)
        
    def add_items(self):
        # First label - button group
        self.add_label_button_group(self.main_window,
                                    button_command=lambda: self.get_file_path(self.get_file_label_text),
                                    button_text='...',
                                    button_x=X_POSITION,
                                    button_y=Y_POSITION + self.group_count * NEW_LINE_HEIGHT + 1,
                                    label_text_var=self.get_file_label_text,
                                    label_text=DEFAULT_GET_FILE_TEXT,
                                    label_x=X_POSITION+BUTTON_WIDTH+BUTTON_LABEL_SPACE,
                                    label_y=Y_POSITION + self.group_count * NEW_LINE_HEIGHT)

        # Second label - button group
        self.add_label_button_group(self.main_window,
                                    button_command=lambda: self.get_export_path(self.get_directory_label_text),
                                    button_text='...',
                                    button_x=X_POSITION,
                                    button_y=Y_POSITION + self.group_count * NEW_LINE_HEIGHT + 1,
                                    label_text_var=self.get_directory_label_text,
                                    label_text=DEFAULT_GET_DIRECTORY_TEXT,
                                    label_x=X_POSITION+BUTTON_WIDTH+BUTTON_LABEL_SPACE,
                                    label_y=Y_POSITION + self.group_count * NEW_LINE_HEIGHT)
     
        convert_button = Button(self.main_window,
                                command=lambda: self.start(self.get_file_label_text.get(),
                                                           self.get_directory_label_text.get()),
                                text='Convert')
        convert_button.pack()
        convert_button.place(x=X_POSITION + (WINDOW_WIDTH - PROGRESS_BAR_WIDTH) / 2,
                             y=Y_POSITION + self.group_count * NEW_LINE_HEIGHT + 1)
        self.group_count += 1
        
    def add_label_button_group(self, window,
                               button_command=None, button_text='', button_x=None, button_y=None,
                               label_text_var=None, label_text='', label_x=None, label_y=None):
        
        button = Button(window, command=button_command, text=button_text)
        button.pack()
        button.place(x=button_x, y=button_y, width=BUTTON_WIDTH, height=BUTTON_HEIGHT)

        label = Label(window, textvariable=label_text_var, **LABEL_SETTINGS)
        label_text_var.set(label_text)
        label.pack(fill=X)
        label.place(x=label_x, y=label_y)
       
        self.group_count += 1
        
        return button, label

    def _paths_valid(self, source, destination):
        if source == DEFAULT_GET_FILE_TEXT:
            self._warning('Source file not found', 'Please select the source file')
            return False
        elif destination == DEFAULT_GET_DIRECTORY_TEXT:
            self._warning('Destination path not found', 'Please select the path to extract file')
            return False

        return True
        
    def _start_progress_bar(self):
        self._frame = ttk.Frame()
        self._progress_bar = ttk.Progressbar(self._frame, length=PROGRESS_BAR_WIDTH, mode='indeterminate')
        self._frame.pack()
        self._progress_bar.pack()
        self._progress_bar.start(25)
        self._frame.place(x=(WINDOW_WIDTH - PROGRESS_BAR_WIDTH) / 2,
                          y=Y_POSITION + self.group_count * NEW_LINE_HEIGHT + 1)

    def _stop_progress_bar(self):
        self._progress_bar.stop()
        
        self._frame.pack_forget()
        self._frame.place_forget()

    def start(self, source, destination):
        if self._paths_valid(source, destination):
            threading.Thread(target=self.run, args=(source, destination,)).start()
            self._start_progress_bar()

    def run(self, source, destination):
        time.sleep(0.1)
        title, message = self.convert(source, destination)

        self._stop_progress_bar()
        self._warning(title or 'Success', message or 'Operation successfully.')
        self.main_window.quit()

    @staticmethod
    def convert(source, destination):
        pass
