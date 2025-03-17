from pdf2docx import Converter
import sys





if __name__ == "__main__":


    usage = """

    =========================================
    Resume Converter: 

    Converts PDF --> Docx  only. 


    usage: python converter.py <resumePDF>

    note: please place the resume
    in the same directory...

    ========================================

    """

    if len(sys.argv) != 2: 
        print(usage)
        exit(0)

    pdf_path = sys.argv[1]
    
    docx_path = "./converted_resume.docx"

    # Acutally try and convert the file to a docx.
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()

    
