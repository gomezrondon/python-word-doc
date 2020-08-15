import docx  #hay que usar es python-docx y se invoca con docx
import datetime


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
     doc = docx.Document('D:\gomez\Documents\python-doc.docx')
     paragraphs  = doc.paragraphs
     text = paragraphs[1].text
     paragraphs[1]._p.clear()

     paragraphs[1].add_run('---------------------> '+str(datetime.datetime.now().today().strftime('%Y-%m-%d  time: %H:%M:%S'))+' \n')
     paragraphs[1].add_run('\n')
     paragraphs[1].add_run(' Nueva entrada de texto!! \n')
     paragraphs[1].add_run('\n')
     paragraphs[1].add_run('\n')
     paragraphs[1].add_run('\n')
     paragraphs[1].add_run('\n')
     #paragraphs[1].add_run('-----------------------------------< \n')
     paragraphs[1].add_run(text)
     doc.save('D:\gomez\Documents\python-doc.docx')

     # print(doc.paragraphs) #[<docx.text.paragraph.Paragraph object at 0x033C6910>, ...
     # for parag in doc.paragraphs:
     #      print(parag.text)

     # doc2 = docx.Document()
     # paraObject = doc2.add_paragraph('this is the matrix')
     # paraObject.add_run(' more text')
     # doc2.save('new-doc.docx')

