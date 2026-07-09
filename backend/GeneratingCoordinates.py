import fitz
import os
import pandas as pd
import re

# this is the main one and it was always used and  not alternate file

folder_path = "../CAS-Original-PDFs"

arr = []
for name in os.listdir(folder_path):
    if name in os.listdir("../backplotted-58"):
        full_path = os.path.join(folder_path, name)
        dr_name = name.split(".")[0]    
        path = full_path
        doc = fitz.open(path)
        page = doc[0]
        page_coordinates = page.rect
        width = page_coordinates.width
        height = page_coordinates.height
        print(dr_name)
        count = 1
        restricted_contents = ["\"","-",".","°","/","\\",",","="]
        pattern = r'^[A-Z]\d+[A-Z]\d+$'
        #"""
        words = page.get_text("words")
        for w in words:
                x1, y1, x2, y2, word, block_no, line_no, word_no = w
                if (not any(char in word for char in restricted_contents)
                    and (48<=ord(min(word))<=57)
                    and (65<=ord(max(word))<=ord('Z'))
                    and (not re.match(pattern, word))
                    and (dr_name[0:-3].upper() not in word)
                    ):
                    data = {
                        'P&ID DRG NO.':dr_name,
                        'TEXT':word,
                        'X1':x1-10,
                        'Y1':y1-10,
                        'X2':x2+10,
                        'Y2':y2+10
                    }
                    arr.append(data)
                    print(word, (x1, y1, x2, y2))

                if (not any(char in word for char in restricted_contents)
                    and len(word) < 3
                    and (dr_name[0:-3].upper() not in word)
                    ):
                    data = {
                        'P&ID DRG NO.':dr_name,
                        'TEXT':word,
                        'X1':x1-10,
                        'Y1':y1-10,
                        'X2':x2+10,
                        'Y2':y2+10
                    }
                    arr.append(data)
                    print(word, (x1, y1, x2, y2))
        #"""

print("reached here")
df_words = pd.DataFrame(arr)
df_words.to_excel("word-data-test-testing.xlsx",sheet_name="sheet1",index=False)
print("Excel generated")




