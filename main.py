from PyPDF2 import PdfFileMerger
from docx2pdf import convert
from pathlib import Path
from win32com import client
import os

class FileConvertor:
    def __init__(self):
        self.command_type = ["wordtopdf","mergepdf","exceltopdf"]
        self.user_command = ""
        self.destination_path = ""

    def word_to_pdf(self,file_name):
        final_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', file_name)
        convert(final_path)

    def excel_to_pdf(self,file_name):
        final_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', file_name)
        excel = client.Dispatch("Excel.Application")
        sheets = excel.Workbooks.Open(final_path)
        work_sheets = sheets.Worksheets[0]
        work_sheets.ExportAsFixedFormat(0, os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', "convertedExcelSheet"))
        sheets.Close(True)

    def merge_pdf(self,folder_path):
        folder_dir = (Path.home()/ "Desktop"/ folder_path)
        pdf_merger = PdfFileMerger(strict=False)
        sorted_files_list = list(folder_dir.glob("*.pdf"))
        sorted_files_list.sort()
        for file in sorted_files_list:
            pdf_merger.append(str(file))
        with Path(Path.home()/"Desktop"/"MergedFiles.pdf").open(mode = "wb") as output_file:
            pdf_merger.write(output_file)

    def obtain_user_command(self):
        while True:
            print("COMMANDS AVAILABLE:")
            for keys in self.command_type:
                print(keys)
            user_request = input("Please enter your request below\n").lower()
            if user_request in self.command_type:
                break
            else:
                print("\nInvalid command please refer to available commands shown below\n")
                continue
        self.user_command = user_request

    def obtain_path(self):
        while True:
            if(self.user_command == "mergepdf"):
                path = input("\nPlease enter a folder/directory to merge pdf files contained in the folder from the desktop\n")
                if os.path.isdir(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop',path)):
                    self.merge_pdf(path)
                    break
            elif(self.user_command == "wordtopdf"):
                file = input("\nPlease enter the word file name with the file extension EX: testing.docx\n")
                if os.path.isfile(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop',file)) and (".docx" in file):
                    self.word_to_pdf(file)
                    break
            elif (self.user_command == "exceltopdf"):
                file = input("\nPlease enter the excel file name with the file extension EX: testing.xlsx\n")
                if os.path.isfile(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', file)) and (".xlsx" in file):
                    self.excel_to_pdf(file)
                    break

            print("\n\nInvalid file/folder provided, please try again")

def main():
    file_conv_object = FileConvertor()
    file_conv_object.obtain_user_command()
    file_conv_object.obtain_path()

if __name__ == "__main__":
    main()





