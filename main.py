import docx  #hay que usar es python-docx y se invoca con docx
import datetime
import time
from open_close_program import ifOpenThenClose, isOpen, findWindowProgram, clickAndSave


def getConfig(f_path):
     config = []
     with open(f_path, "rt") as t_file:
          for each in t_file:
               config.append(each.replace("\n",""))
     return config

def getTextFromFile(f_path):
     lines = []
     with open(f_path, "rt") as t_file:
          for each in t_file:
               lines.append(each)
     return lines


def writeTextToWord(paste_text):
     for line in paste_text:
          paragraphs[0].add_run(line)
     paragraphs[0].add_run('\n')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
     config = getConfig('config.txt')
     filePath = config[0]
     docName = config[1]

     if isOpen(docName) == True:
          x, y = findWindowProgram(docName)
          clickAndSave(x, y)
          ifOpenThenClose(filePath, docName)
          time.sleep(1)


     doc = docx.Document(filePath)
     paragraphs  = doc.paragraphs
     text = paragraphs[0].text
     paragraphs[0]._p.clear()
     paste_text = getTextFromFile(config[2])
     t_paste_text = getTextFromFile(config[3])
     # image_path = r'D:\temp\oct-tesseract\image\paste.png'

     paragraphs[0].add_run('------------1---------> '+str(datetime.datetime.now().today().strftime('%Y-%m-%d  time: %H:%M:%S'))+' \n')
     paragraphs[0].add_run('\n')

     if len(paste_text) > 0:
          paragraphs[0].add_run(' ** Original text ** \n')
          writeTextToWord(paste_text)
          paragraphs[0].add_run(' ** Translated text ** \n')
          writeTextToWord(t_paste_text)
     else:
          paragraphs[0].add_run(' Nueva entrada de texto!! \n')
          paragraphs[0].add_run('\n')
          paragraphs[0].add_run('\n')
          paragraphs[0].add_run('\n')

     paragraphs[0].add_run('\n')

     paragraphs[0].add_run(text)
     doc.save(filePath)

     if isOpen(docName) == False:
          ifOpenThenClose(filePath,docName)
