import os
import csv
from rutermextract import TermExtractor
term_extractor = TermExtractor()
filename = input("Введите путь к файлу: ")
if not os.path.exists(filename):
    print("Указанный файл не существует")
else:
     with open(filename, encoding="utf8") as file:
         text = file.read()
         for term in term_extractor(text):
             print(term.normalized.capitalize(), term.count)
         
     with open('example2.csv', 'w+', encoding='utf8', newline='') as out_file: 
            writer=csv.writer(out_file,delimiter=';')
            for term in term_extractor(text):
                writer.writerows(term.normalized.capitalize())
     print('Запись .csv-файла завершена.')
