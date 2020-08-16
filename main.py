import docx  #hay que usar es python-docx y se invoca con docx
import datetime
import time
from open_close_program import ifOpenThenClose, isOpen, findWindowProgram, clickAndSave

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
     filePath = r'C:\Users\jrgm\Documents\GAA\F603.docx'
     proc_name = 'WINWORD.EXE'
     if isOpen(proc_name) == True:
          docx = 'F603.docx'
          x, y = findWindowProgram(docx)
          clickAndSave(x, y)
          ifOpenThenClose(filePath,proc_name)
          time.sleep(1)


     doc = docx.Document(filePath)
     paragraphs  = doc.paragraphs
     text = paragraphs[0].text
     paragraphs[0]._p.clear()

     paragraphs[0].add_run('------------1---------> '+str(datetime.datetime.now().today().strftime('%Y-%m-%d  time: %H:%M:%S'))+' \n')
     paragraphs[0].add_run('\n')
     paragraphs[0].add_run(' Nueva entrada de texto!! \n')
     paragraphs[0].add_run('\n')
     paragraphs[0].add_run('\n')
     paragraphs[0].add_run('\n')
     paragraphs[0].add_run('\n')
     #paragraphs[1].add_run('-----------------------------------< \n')
     paragraphs[0].add_run(text)
     doc.save(filePath)

     if isOpen(proc_name) == False:
          ifOpenThenClose(filePath,proc_name)
