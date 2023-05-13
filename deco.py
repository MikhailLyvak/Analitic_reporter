import time


filename1 = "/home/izevsl/Документы/Analitic_reporter/1C_Repo/01Klientu.xlsx"
filename2 = "/home/izevsl/Документы/Analitic_reporter/1C_Repo/02Продажи.xlsx"

# Декоратор який виводить інформацію якщо файл зчитано
def file_parsed(func):
    
    def inner(*args, **kwargs):
        res = func(*args, **kwargs)
        print(f"Інформацію успішно зчитано з Excel файлу --> {''.join(filename2.split('/')[-1:])}")
        return res
    return inner


# Декоратор який заміряє час виконання функції (написано для зчитки ексель файлів)
def parsing_time(func):
    
    def inner(*args, **kwargs):
        start = time.time()
        res = func(*args, **kwargs)
        end = time.time()
        print(f"Час зчитування файлу склав {round(end - start, 2)} сек.")
        return res
    
    return inner
