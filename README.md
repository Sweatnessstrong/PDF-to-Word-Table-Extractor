This project is to extract table from Portable Document File(PDF) and store one to Word file.
I use the camelot library for this process.

# PDF to Word Table Extractor

This Python script extracts tables from a PDF file and creates a Word document with the extracted tables.

## Requirements

- Python 3.x
- `camelot-py` library
- `python-docx` library

You can install the required libraries using pip:

```bash
pip install camelot-py python-docx
```

## Usage

1. Place the PDF file with tables in the same directory as the script.
2. Update the `pdf_file` variable with the name of your PDF file.
3. Run the script using Python:

```bash
python pdf_to_word.py
```

4. The script will create a Word document named `Project_Complexity_Classification_Guide2.docx` in the same directory, containing the extracted tables.

## How it Works

1. The script specifies the path to the PDF file with tables using the `pdf_file` variable.
2. It creates a new Word document using the `Document()` function from `python-docx`.
3. The `camelot.read_pdf()` function is used to extract tables from the PDF file. The `pages='all'` parameter extracts tables from all pages of the PDF.
4. The `flavor='stream'` parameter is used to parse tables that have white spaces between cells, simulating a table structure.
5. The `layout_kwargs` parameter is used to adjust the table detection settings. `{'detect_vertical': True, 'all_texts': True}` detects vertical text and includes all text in the table.
6. The script iterates through each extracted table and adds it to the Word document using the `add_table()` function from `python-docx`.
7. Finally, the Word document is saved to the specified file name using the `save()` function.

## Customization

You can customize the script by:

- Changing the `pdf_file` variable to specify a different PDF file.
- Adjusting the `layout_kwargs` settings to improve table detection for your specific PDF file.
- Modifying the output file name by changing the argument passed to the `save()` function.

## Dependencies

- `camelot-py`: A Python library for extracting tables from PDF files.
- `python-docx`: A Python library for creating and manipulating Word documents.

