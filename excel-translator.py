import os
import openpyxl
from deep_translator import GoogleTranslator

#start read excel
for filename in os.listdir():
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        print(filename)
        #start translate
        workbook = openpyxl.load_workbook(filename)
        translator = GoogleTranslator(source="ja", target="en")

        for sheet in workbook.sheetnames:
            worksheet = workbook[sheet]
            
            for row in worksheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        try:
                            cell.value = translator.translate(cell.value)
                        except Exception as e:
                            print(f"Error translating cell {cell.coordinate}: {e}")

        translated_filename = f"translated_{filename}"
        workbook.save(translated_filename)

        print(f"Translation completed and saved to {translated_filename}")
