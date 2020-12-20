import openpyexcel as xl
from time import sleep

class Excel_Joiner:
    def __init__(self, main_file_location):
        self.wb = xl.load_workbook(main_file_location)
        self.wbs = self.wb.active

        self.getting_index_of_max_row_of_wb1s = self.wbs.cell(
            self.wbs.max_row, 1).value
        self.record_count = self.getting_index_of_max_row_of_wb1s

        # value of 1st column
        self.first_record_in_wbs = [self.wbs.cell(
            1, main_cell).value for main_cell in range(1, self.wbs.max_column+1)]

    def add_sheet(self, sheet_location):

        wb2 = xl.load_workbook(sheet_location)
        wb2s = wb2.active

        for row in wb2s.iter_rows():

            # getting the record to a list
            record_in_wb2s = [cell.value for cell in row]
            
            maxrow = self.wbs.max_row

            # passing the 1st row
            if record_in_wb2s == self.first_record_in_wbs:
                pass
            else:
                self.record_count = self.record_count+1

                # for cell in range(1, self.wbs.max_column+1):
                #     new_cell = self.wbs.cell(maxrow+1, cell)
                #     new_cell.value = record_in_wb2s[cell-1]

                # indexing if needed REMEBER TO DISABLE THE ABOVE FOR LOOP

                indexing_cell = self.wbs.cell(maxrow+1, 1)
                indexing_cell.value = self.record_count

                for cell in range(2, self.wbs.max_column+1):
                    new_cell = self.wbs.cell(maxrow+1, cell)
                    new_cell.value = record_in_wb2s[cell-1]

                if self.record_count % 1000 == 0:
                    print(self.record_count-self.getting_index_of_max_row_of_wb1s,
                          " Records inserted \nplease wait...")

        print(self.record_count, " Rows Available \nsucess!!!")

    def save(self):

        try:
            while True:
                self.wb.save('Excel_File.xlsx')
                print("file saved as 'Excel_File.xlsx'")
                sleep(16)
                break
            
        except:
            print("This File is \"opened in another program\". Please close it to save !")

       


xljoin = Excel_Joiner("file_1.xlsx")
xljoin.add_sheet("file_2.xlsx")
xljoin.add_sheet("file_2.xlsx")
xljoin.save()
