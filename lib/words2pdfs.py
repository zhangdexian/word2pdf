#coding=utf-8
import os,shutil,sys
from win32com import client
import time

success = 0
fail = 0

def encode_content(content=''):
    return content.decode('utf-8').encode('cp936')

def doc2pdf(doc_name, pdf_name):
    """
    :word文件转pdf
    :param doc_name word文件名称
    :param pdf_name 转换后pdf文件名称
    """
    global success, fail
    try:
        word = client.DispatchEx("Word.Application")
        if os.path.exists(pdf_name):
            os.remove(pdf_name)
        worddoc = word.Documents.Open(doc_name,ReadOnly = 1)
        worddoc.SaveAs(pdf_name, FileFormat = 17)
        success += 1
        print '%s >>>>>>> %s   转换成功'.decode('utf-8').encode('cp936') %(os.path.split(doc_name)[1] , os.path.split(pdf_name)[1])
        return pdf_name
    except Exception as e:
        fail += 1
        print(e)
        print encode_content('%s 转换失败..' % doc_name)
        return 1
    finally:
        worddoc.Close()
        word.Quit()

def doc2docx(doc_name,docx_name):
    """
    :doc转docx
    """
    try:
        # 首先将doc转换成docx
        word = client.Dispatch("Word.Application")
        doc = word.Documents.Open(doc_name)
        #使用参数16表示将doc转换成docx
        doc.SaveAs(docx_name,16)
    except:
        pass
    finally:
        doc.Close()
        word.Quit()

def createDirs(basePath=os.getcwd()):
    # 存放转化后的pdf文件夹
    pdfs_dir = basePath + '/pdfs'
    if not os.path.exists(pdfs_dir):
        os.mkdir(pdfs_dir)
    return pdfs_dir

def getFileNames(basePath=os.getcwd()):
    filenames=[]
    # move all .words files to words_dir
    for file in os.listdir(basePath):
        if file.endswith('.docx'):
            filenames.append(file)
        elif file.endswith('.doc'):
            filenames.append(file)
        else:
            pass
    return filenames

def convert(basePath=os.getcwd(),filenames=[]):
    pdfs_dir=createDirs(basePath)
    for filename in filenames:
        pdfName='.'.join(filename.split('.')[:-1])+'.pdf'
        doc2pdf(os.path.join(basePath,filename),os.path.join(pdfs_dir,pdfName))

if __name__ == '__main__':
    basePath=os.getcwd()
    lfileNames=getFileNames(basePath)
    if len(lfileNames) == 0:
        print encode_content('未检测到word文档')
        time.sleep(5)
        sys.exit()
    print encode_content('内容分析中...')
    print encode_content('检测到当前目录下共有word文档：%d' % len(lfileNames))
    print('are you going to convert these files to pdf?')
    for filename in lfileNames:
        print(filename)
    print("input \"yes\" to start convert, and \"no\" to quik program")
    while True:
        command=raw_input()
        if command=='yes':
            convert(basePath,lfileNames)
            print(encode_content('成功： %d' % success))
            print(encode_content('失败： %d' % fail))
            time.sleep(5)
            break
        elif command=='no':
            break
        else:
            print('wrong command,input yes or no please')
