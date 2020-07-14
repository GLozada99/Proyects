import bs4 #for html parsing
import urllib.request as req
from urllib.request  import urlopen #for opening the url
import re
import openpyxl
import yagmail

def correction(s):
    s = re.sub(r'\. ','.',s)
    s = re.sub(r'\.','. ',s)#replacing every period with period and space
    s = re.sub(r'\?','? ',s)#replacing every ? with ? and space
    return s

def get_text_parsed(url):
    url_open =  req.urlopen(url)
    text = url_open.read()
    soup = bs4.BeautifulSoup(text,'html.parser')
    return soup

def fill_dict(soup,data_dict):
    #Getting needed data from page
    data_dict['Phrase'] = '\n' + soup.findAll("div",{"class":"field-item even"})[1].text + '\n'  
    data_dict['Definition'] = '\n' + soup.findAll("div",{"class":"field-item even"})[2].text + '\n'  
    example = soup.findAll("div",{"id":"bootstrap-panel-2-body"})[0].text
    data_dict['Example'] = correction(example)
    return data_dict

def create_load_workbook(filename):
    #Loading workbook and getting first spreadsheet
    #If the file doesn't exists, it is created
    try:
        workbook = openpyxl.load_workbook(filename = filename)#load workbook with the name
        sheet = workbook['Phrases']
    except:
        workbook = openpyxl.Workbook()
        elim = workbook['Sheet']
        workbook.remove_sheet(elim)
        sheet = workbook.create_sheet('Phrases')
        for i in range(3):
            print(1,i+1,headers[i])
            sheet.cell(1,i+1).value = headers[i]
            sheet.cell(1,i+1).font = openpyxl.styles.Font(bold=True) 
    return (workbook,sheet)
        
def get_emails(file):
    fil = open(file,'r')
    name_ls = []
    mail_ls = []
    dicti = {'Names':name_ls,'Mails':mail_ls}
    for line in fil:
        stripped_line = line.strip().split(" ")
        dicti['Names'].append(stripped_line[0])
        dicti['Mails'].append(stripped_line[1])
    return dicti

def send_mails(emitter,contacts,contenct):
    yag = yagmail.SMTP(emitter)
    phr = contenct['Phrase']
    des = contenct['Definition']
    exa = contenct['Example']
    body = 'Phrase: {}\nDefinition: {}\nExample: {}'.format(phr,des,exa) 
    for name,reciever in zip(contacts['Names'],contacts['Mails']):
        yag.send(
            to=reciever,
            subject="Hola {}, esta es la frase de hoy".format(name),
            contents=body, 
        )
    yag.close()
    
            
#Connecting to url and constructing dict for the data
my_url = 'https://www.ihbristol.com/english-phrases'
soup = get_text_parsed(my_url)
headers = ['Phrase','Definition','Example']
data_dict = {headers[0]:'',headers[1]:'',headers[2]:''}
data_dict = fill_dict(soup,data_dict)

excel_filename = '/home/gustavolozada/Documents/PyPrograming/ExcelDocs/PhrasesOfTheDay.xlsx'
workbook, sheet = create_load_workbook(excel_filename) 

#Iterating through the dict to write on spreadsheet
blank_row = sheet.max_row
for head,col in zip(headers,range(3)):
    print(head + ": "+ data_dict[head])

    #The row and column must be at least 1
    cell = (blank_row + 1, col + 1,openpyxl.utils.get_column_letter(col + 1)) #row,column,column_letter

    sheet.cell(cell[0],cell[1]).value = data_dict[head] 
    
    sheet.cell(cell[0],cell[1]).alignment = openpyxl.styles.Alignment\
    (wrapText=True, horizontal='center',vertical='center')
    
    sheet.column_dimensions[cell[2]].width = float(20 + 15*col)

sheet.row_dimensions[blank_row+1].height = float(60)#last column (description) must have bigger width

#Save and close the excel file
workbook.save(excel_filename)
workbook.close()

#Get and open contacts file
contacts_file = '/home/gustavolozada/Documents/PyPrograming/Contacts.txt'
contacts = get_emails(contacts_file)

emitter = 'gu.lozada9.mail@gmail.com'
send_mails(emitter,contacts,data_dict)
