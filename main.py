from tkinter_demo import ExcelConverterGui


class ExcelConverterGuiOverride(ExcelConverterGui):
    @staticmethod
    def convert(source, destination):
        print(source, destination)
        title = 'Success'
        message = 'File successfully converted'

        return title, message


gui = ExcelConverterGuiOverride()
gui.start_gui()
