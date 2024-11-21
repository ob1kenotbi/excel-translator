import openpyxl
from deep_translator import GoogleTranslator

workbook = openpyxl.load_workbook('1.xlsx')
translator = GoogleTranslator(source='ja', target='en')

for sheet in workbook.sheetnames:
    worksheet = workbook[sheet]
    
    for row in worksheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                try:
                    cell.value = translator.translate(cell.value)
                except Exception as e:
                    print(f"Error translating cell {cell.coordinate}: {e}")

workbook.save('translated_1.xlsx')

print("Translation completed and saved to 'translated_1.xlsx'")
