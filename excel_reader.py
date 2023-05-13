import openpyxl
from deco import parsing_time, filename1, filename2
from file_mover import file_mover
# from GUI import filename1
# from GUI import filename2


@parsing_time
def excel_reader(vertical_start: int, gorizontal_start: int, gorizontal_end: int, file_place: str) -> list:
    if file_place:
        book = openpyxl.open(file_place)
        sheet = book.active
        
        ListRow1 = [
            list(sheet[row][col].value for col in range(gorizontal_start, gorizontal_end)) 
            for row in range(vertical_start, sheet.max_row+1) if sheet.row_dimensions[row].hidden == False
        ]
        
        # for i in ListRow1:
        #     print(i)
        
        return ListRow1
    print("File was not provided")

# excel_reader(7, 1, 9, filename1)
# excel_reader(7, 1, 7, filename2)
