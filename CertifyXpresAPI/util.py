from docx import Document
from docx.shared import Pt
import os
import win32com.client
import schema
from datetime import datetime

def get_run_formatting(run):
    """
    Extract formatting information from a Run object.
    """
    formatting = {
        'font': run.font.name,
        'size': run.font.size,
        'color': run.font.color.rgb,
        'bold': run.font.bold,
        'italic': run.font.italic,
        'underline': run.font.underline,
    }
    return formatting


def apply_formatting_to_run(run, formatting):
    """
    Apply formatting to a Run object.
    """
    run.font.name = formatting['font']
    run.font.size = formatting['size']
    run.font.color.rgb = formatting['color']
    run.font.bold = formatting['bold']
    run.font.italic = formatting['italic']
    run.font.underline = formatting['underline']


async def convert_to_pdf(docx_path, pdf_path):
    # Open Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Don't show the Word application window

    # Open the document
    doc = word.Documents.Open(docx_path)

    # Save as PDF
    doc.SaveAs(pdf_path, FileFormat=17)  # FileFormat=17 for PDF

    # Close the document and Word application
    doc.Close()
    word.Quit()



# Multiple
async def replace_text_in_docx(template_path, certificate, output_path):
    # Load the template document
    doc = Document(template_path)
    now = datetime.now()

    # Defined a dictionary of words to replace and their replacements
    replacements = {
        # Certificate Holder Name
        'person': certificate.name,
        # certification type
        'header': certificate.certification_type,
        # achievement
        'achtype': certificate.achivement_type,
        # mentor
        'mentid': certificate.mentor_name,
        '[mentor]': certificate.mentor_type,
        # institute
        'institute': certificate.insititue_name,
        # course
        'course': certificate.course_name,
        # sign of mentor
        'Msign': certificate.mentor_name,

        # director 
        'mentor2' : certificate.director_name,
        'menpos2' : certificate.director_type,
        'Osign' : certificate.director_name,
        
        # date
        'date': str(now.day),
        'month': str(now.strftime('%B')),
        'year': str(now.year)
    }
        
    print('replacements:', replacements)

    # Replace text in all paragraphs
    for para in doc.paragraphs:
        for run in para.runs:
            words = run.text.split()
            for i, word in enumerate(words):
                if word in replacements.keys():
                    formatting = get_run_formatting(run)
                    words[i] = replacements[word]
                    run.text = ' '.join(words)
                    apply_formatting_to_run(run, formatting)
                    print(f"{word} replaced")


    # Save the modified document
    doc.save(output_path)