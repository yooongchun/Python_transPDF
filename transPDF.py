# 该程序实现读取word文件中的文字内容并按照特定的规则替换文字

# -*- encoding: utf8 -*-

__author__ = 'yooongchun'

# 引入所需要的基本包
import os
import re
import xlrd
import win32com.client
import logging
logging.basicConfig(level=logging.INFO)


# 处理Word文档的类
class RemoteWord:
    def __init__(self, filename=None):
        self.xlApp=win32com.client.DispatchEx('Word.Application')
        self.xlApp.Visible=0
        self.xlApp.DisplayAlerts=0    #后台运行，不显示，不警告
        if filename:
            self.filename=filename
            if os.path.exists(self.filename):
                self.doc=self.xlApp.Documents.Open(filename)
            else:
                self.doc = self.xlApp.Documents.Add()    #创建新的文档
                self.doc.SaveAs(filename)
        else:
            self.doc=self.xlApp.Documents.Add()
            self.filename=''

    def add_doc_end(self, string):
        '''在文档末尾添加内容'''
        rangee = self.doc.Range()
        rangee.InsertAfter('\n'+string)

    def add_doc_start(self, string):
        '''在文档开头添加内容'''
        rangee = self.doc.Range(0, 0)
        rangee.InsertBefore(string+'\n')

    def insert_doc(self, insertPos, string):
        '''在文档insertPos位置添加内容'''
        rangee = self.doc.Range(0, insertPos)
        if (insertPos == 0):
            rangee.InsertAfter(string)
        else:
            rangee.InsertAfter('\n'+string)

    def replace_doc(self,string,new_string):
        # self.doc.Content.Find.Execute(FindText=string,ReplaceWith=new_string,Replace=2,Wrap=1)
        self.xlApp.Selection.Find.ClearFormatting()
        self.xlApp.Selection.Find.Replacement.ClearFormatting()
        self.xlApp.Selection.Find.Execute(string, False, False, False, False, False, True, 1, True, new_string, 2)

    def save(self):
        '''保存文档'''
        self.doc.Save()

    def save_as(self, filename):
        '''文档另存为'''
        self.doc.SaveAs(filename)

    def close(self):
        '''保存文件、关闭文件'''
        self.save()
        self.xlApp.Documents.Close()
        self.xlApp.Quit()


# 遍历找到word文件路径
def find_docx(pdf_path):
    file_list=[]
    if os.path.isfile(pdf_path):
        file_list.append(pdf_path)
    else:
        for top, dirs, files in os.walk(pdf_path):
            for filename in files:
                if filename.endswith('.docx')or filename.endswith('.doc'):
                    abspath = os.path.join(top, filename)
                    file_list.append(abspath)
    return file_list


# 替换文本内容
def replace_docx(rule,docx_list):
    len_doc=len(docx_list)
    i=0  # 计数
    for docx in docx_list:
        i+=1
        logging.info('开始替换第 %s/%s 个word文件内容:%s...'%(i,len_doc,os.path.basename(docx)))
        doc = RemoteWord(docx)  # 初始化一个doc对象
        for item in rule:  # 替换
            doc.replace_doc(item[0], item[1])
        doc.close()

    logging.info('完成！')


# 对内容进行排序
# 这里因为在进行文本替换的时候涉及到一个长句里面的部分可能被短句（相同内容）内容替换掉
# 因而必须先把文本按照从长到短的顺序来进行替换
def sort_rule(rule):
    result=[]
    for item, val in rule.items():
        le=len(item)
        flag = True
        if len(result)>0:
            for index, res in enumerate(result):
                if len(item) >= len(res[0]):
                    flag=False
                    result.insert(index, (item, val))
                    break
            if flag:
                result.append((item, val))

        else:
            result.append((item,val))

    return result


# 加载Excel,把取得的内容返回，格式：dict{'原文':'译文'}
def init_excel(excel_path):
    logging.info('加载文本匹配规则的Excel:%s' % os.path.basename(excel_path))
    rule={}  # 储存原文和翻译内容
    pdf_path=''
    try:
        book = xlrd.open_workbook(excel_path)  # 打开一个wordbook
        sheet = book.sheet_by_name('Translation')  # 切换sheet
        rows = sheet.nrows  # 行数
        for row in range(rows - 1):
            text_ori=sheet.cell(row, 0).value  # 取得数据：原文
            text_trans=sheet.cell(row,1).value  # 取得数据：译文
            if not re.match(r'^#.+',text_ori):  # 原文不以#开头
                if text_ori == 'pdf文件(或文件夹)地址':   # 获得pdf文件路径
                    pdf_path=text_trans
                else:
                    rule[text_ori]=text_trans  # 取得值加入text
    except IOError:
        raise IOError
    logging.info('加载Excel完成！')

    return pdf_path, rule

if __name__ == '__main__':

    excel_path = './match_rule.xlsx'    # 替换规则的Excel文件地址
    logging.info('正在打开pdf转换软件，请手动转换你的pdf文件！')
    os.popen(r'"./PDF2Word/pdf2word.exe"')
    flag=input('你已经完成pdf文件转换了吗？(y/n)：')
    while not flag == 'y':
        logging.info('请先转换pdf！')
        flag = input('你已经完成pdf文件转换了吗？(y/n)：')
    pdf_path, rule = init_excel(excel_path)  # 加载Excel,取得内容
    sorted_rule=sort_rule(rule)  # 排序规则：按照由长到短
    docx_list=find_docx(pdf_path)  # 获取docx文件路径
    replace_docx(sorted_rule,docx_list)  # 替换内容

    logging.info('程序执行完成!')
