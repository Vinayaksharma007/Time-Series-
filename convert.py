import comtypes.client
import os

def convert_doc_to_docx(doc_path):
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(doc_path)
    docx_path = os.path.splitext(doc_path)[0] + ".docx"
    
    # Save as .docx
    doc.SaveAs(docx_path, FileFormat=16)  # 16 is the file format for .docx
    
    doc.Close()
    word.Quit()
    
    if os.path.exists(docx_path):
        print(f"Successfully converted '{doc_path}' to '{docx_path}'")
        return docx_path
    else:
        print(f"Conversion failed, .docx file not found.")
        return None

# Example usage
doc_file_path = r"C:\Users\sharm\OneDrive\Desktop\Scrapper\New folder\RP-172476  TS 38.101-2 Presentation for RAN#78.doc"  # Replace with your .doc file path
converted_docx_path = convert_doc_to_docx(doc_file_path)

if converted_docx_path:
    print(f"File converted and saved at: {converted_docx_path}")
else:
    print("File conversion failed.")