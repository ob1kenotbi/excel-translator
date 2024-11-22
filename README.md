# Excel File Translator

This Python program translates all Excel files containing Japanese words in a specified directory into English. After translation, the program organizes the files into respective folders: `Archive` for the original files and `Translate` for the translated versions.

---

## Features
- Automatically detects Excel files in the directory.
- Scans for Japanese words and translates them into English.
- Moves the original Excel files to the `Archive` folder.
- Saves the translated Excel files in the `Translate` folder.

---

## Prerequisites
1. **Python**: Ensure you have Python installed (preferably Python 3.8 or later).
2. **Libraries**: Install the following Python libraries:
   - `openpyxl`: For reading and writing `.xlsx` files.
   - `deep-translator`: For translating Japanese to English.
