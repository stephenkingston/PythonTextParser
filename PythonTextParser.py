from bs4 import BeautifulSoup
import re
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt
from os import path

answerStart = 'class="cls_009">Solution:</span></div>'
answerEnd = '<span class="cls_007">'
answerContentDiv1 = "cls_010"
answerContentDiv2 = "cls_004"

questionStart = '<div class="cls_007"'
questionEnd = 'class="cls_009">Solution:</span></div>'
questionContentDiv = "cls_007"

currentDir = path.dirname(path.realpath(__file__))

with open(currentDir + "/source/Test_source.html", 'rb', ) as html:
    soup = BeautifulSoup(html, features="lxml")

soup_string = str(soup)
new_soup_string = soup_string.replace("</body></html>",
                                      '<div style="position:absolute;left:54.00px;top:679.54px" '
                                      'class="cls_007">{}</body></html>'.format(answerEnd))


questionPattern = re.findall(r'(?m){}.*?'
                             r'{}'.format(questionStart, questionEnd), new_soup_string, flags=re.S)
answerPattern = re.findall(r'(?m){}.*?'
                           r'{}'.format(answerStart, answerEnd), new_soup_string, flags=re.S)

ExcelQuestions = []
ExcelAnswers = []

for question in questionPattern:
    soup2 = BeautifulSoup(question, features="lxml")
    question_parsed = soup2.findAll("div", {"class": questionContentDiv})
    question = ''
    for q in question_parsed:
        question = question + q.text + '\n'
    ExcelQuestions.append(question)

for answer in answerPattern:
    soup3 = BeautifulSoup(answer, features="lxml")
    answer_parsed = soup3.findAll("div", {"class": answerContentDiv1})
    answer_parsed2 = soup3.findAll("div", {"class": answerContentDiv2})
    answer = ''
    answer2 = ''
    for a in answer_parsed:
        answer = answer + a.text + '\n'
    ExcelAnswers.append(answer)

    for a2 in answer_parsed2:
        answer2 = answer2 + a2.text + '\n'
    ExcelAnswers.append(answer2)

while '' in ExcelAnswers:
    ExcelAnswers.remove('')

while '' in ExcelQuestions:
    ExcelQuestions.remove('')

file = open_workbook(currentDir + "/output/output_format.xls")

sheet1 = copy(file)
sheet = sheet1.get_sheet(0)
style = xlwt.XFStyle()
style.alignment.wrap = 1

for i in range(0, len(ExcelQuestions)):
    sheet.write(1+i, 0, ExcelQuestions[i], style)

for j in range(0, len(ExcelAnswers)):
    sheet.write(1+j, 1, ExcelAnswers[j], style)

sheet1.save("output/Parsed.xls")
