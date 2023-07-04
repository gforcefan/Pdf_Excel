import pdfplumber
from openpyxl import Workbook


class PDF(object):
    def __init__(self, file_path):
        self.pdf_path = file_path
        try:
            self.pdf_info = pdfplumber.open(self.pdf_path)
            print('Read done！')
        except Exception as e:
            print('Read filed!', e)

    def get_table(self):
        wb = Workbook()  
        ws = wb.active  
        con = 0
        try:
            for page in self.pdf_info.pages:
                for table in page.extract_tables():
                    for row in table:
                        row_list = [cell.replace('\n', '').replace('\r', '') if cell else '' for cell in row]
                        ws.append(row_list)  
                con += 1
                print('---------------%s page---------------' % con)
        except Exception as e:
            print('ero：', e)
        finally:
            wb.save('\\'.join(self.pdf_path.split('\\')[:-1]) + '\pdf_excel.xlsx')
            print('Writedown')
            self.close_pdf()


    def close_pdf(self):
        self.pdf_info.close()


if __name__ == "__main__":
    file_path = input('')
    pdf_info = PDF(file_path)
    pdf_info.get_table()
