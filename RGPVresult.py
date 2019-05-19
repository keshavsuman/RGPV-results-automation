from selenium import webdriver
from selenium.webdriver.support.select import Select
import pyautogui as p,openpyxl as xl,time


def fill(i):
    roll=w.find_element_by_id("ctl00_ContentPlaceHolder1_txtrollno")
    roll.clear()
    roll.send_keys(sheet.cell(row=i, column=1).value)
    Select(w.find_element_by_id("ctl00_ContentPlaceHolder1_drpSemester")).select_by_index(semester-1)
    # pic=p.screenshot(region=(816,568,250,70))
    # pic.save('try.png')
    # result = pytesseract.image_to_string(pic).upper().replace(" ","")
    text = w.find_element_by_id("ctl00_ContentPlaceHolder1_TextBox1")
    text.clear()
    result=p.prompt('enter the captcha here').upper().replace(' ','')
    text.send_keys(result)
    w.find_element_by_id("ctl00_ContentPlaceHolder1_btnviewresult").click()


def getdata():
    details=[]
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblNameGrading").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblProgramGrading").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblBranchGrading").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblSemesterGrading").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblSGPA").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblcgpa").text)
    writedata(details)

def writedata(details):
    for j in range(2,sheet.max_column+1):
        sheet.cell(row=i,column=j).value=details[j-2]

# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
w = webdriver.Chrome()
w.maximize_window()
w.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
path=p.prompt("Enter the Name of your excel sheet(With full path if not in same folder)")
semester=int(p.prompt("Enter the semester here"))
wb=xl.load_workbook(path)
sheet=wb.active
print('Your script is running...')
sheet.cell(row=1,column=2).value='Name'
sheet.cell(row=1,column=3).value='Course'
sheet.cell(row=1,column=4).value='Branch'
sheet.cell(row=1,column=5).value='Semester'
sheet.cell(row=1,column=6).value='SGPA'
sheet.cell(row=1,column=7).value='CGPA'
print('Excel sheet has been initialised...')
w.find_element_by_id("radlstProgram_1").click()
i=2
while(sheet.cell(row=i,column=1).value):
    fill(i)
    getdata()    
    w.find_element_by_id("ctl00_ContentPlaceHolder1_btnReset").click()
    i=i+1
wb.save(path)
print('Task completed...')
w.close()
#selenium.common.exceptions.UnexpectedAlertPresentException