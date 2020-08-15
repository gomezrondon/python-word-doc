import docx  #hay que usar es python-docx y se invoca con docx



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
     doc = docx.Document('D:\gomez\Documents\python-doc.docx')
     print(doc.paragraphs) #[<docx.text.paragraph.Paragraph object at 0x033C6910>, ...
     for parag in doc.paragraphs:
          print(parag.text)

