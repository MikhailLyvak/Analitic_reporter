import shutil
import os
from pathlib import Path
from datetime import datetime


"""
this function moving all -> files like (excel/text) <-
into folder named -> Звіти за 'сьогоднішня дата' <-
"""


def file_mover() -> None:
    LST = [".xlsx", ".txt"]
    
    now = datetime.now().strftime("%d_%m_%Y")
    folder_name = f"Report_for_{now}"
    Path(folder_name).mkdir(exist_ok=True)
    for val in LST:
        for file in [file for file in os.listdir() if val in file]:
            shutil.move(file, folder_name)
            print(f"-->  {file}  <-- has been moved in folder -->  reports  <--")

    os.rename(folder_name, f'Звіти за {".".join(now.split("_"))}')

