from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def modify_word_document(input_docx, output_docx, position, company, industry):
    # Load the Word document
    doc = Document(input_docx)

    # Iterate through all paragraphs in the document
    for paragraph in doc.paragraphs:
        # Modify specific elements (e.g., replace text)
        if "[POSITION]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[POSITION]", position)
        if "[COMPANY]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[COMPANY]", company)
        if "[INDUSTRY]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[INDUSTRY]", industry)

    # Save the modified document
    doc.save(output_docx)


if __name__ == "__main__":
    input_docx = "Motivation Letter - Design.docx"  # Path to the input Design Word document
    # input_docx = "Motivation Letter - IT.docx"  # Path to the input IT Word document
    output_docx = "/home/barry/workspace/github.com/hbbry/CVs/Motivation_Letter.docx"  # Path to save the modified Word document


    # Modify the Word document
    modify_word_document(input_docx, output_docx, "Body Design Engineer", "Toyota", "Automotive")
