##Lee Kamar
##MVP-pdf to docx Converter
##2/09/2022

import os
import win32com.client

word=win32com.client.Dispatch("word.Application")
word.visible=0

prompt1=input("Do you want to transcribe text from image, then convert document text to docx format?( type 'yes' or 'no'):")
if prompt1=='yes':
    print("opening image processing program")
else:
    print("PDF to DOCX document converter \n"
                "Instructions for use: \n"
                "1. Place pdf file in the same directory as the doc_converter program \n"
                "2. Run the python program and enter the name of the pdf file to be converted to docx \n"
                "3. If successful, converted document will be in the same directory as the original. \n"
                )
    doc_to_convert=input("Please type the name of the pdf document to be converted:")
    input_file=os.path.abspath(doc_to_convert)

    try:
        wb=word.Documents.Open(input_file)
        output_file=os.path.abspath(doc_to_convert[0:-4] + ".docx".format())
        wb.SaveAs2(output_file,FileFormat=16)
        print("Document conversion from pdf to docx is complete!")
        wb.Close()

    except:
        print("An error has occurred. Verify that you have entered the file name correctly and retry")