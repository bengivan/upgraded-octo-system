import csv
from fpdf import FPDF
from PyQt6.QtWidgets import *
import sys
from functools import partial
import os
import numpy as np
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, inch, mm
from reportlab.platypus import Table, LongTable, TableStyle, SimpleDocTemplate, Paragraph, Image, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import *
import datetime
from datetime import date
from dateutil.parser import parse
import pandas as p

title = 'WEEKLY REPORT'
pageinfo = "This document is strictly private and confidential"

def error_msg(error):
    error_win = QApplication([])
    
    error_dialog = QErrorMessage()
    error_dialog.showMessage(error)
    print(error)
    sys.exit(error_win.exec())
    return error_win, error_dialog


def file_explorer(button_id):
    dialog = QFileDialog()
    if button_id == 'csv_browse':
        dialog.setDirectory(r'C:\\Downloads')
        #dialog.setDirectory(r'\downloads')
        dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        dialog.setNameFilter("Text (*.csv)")
        dialog.setViewMode(QFileDialog.ViewMode.List)
        if dialog.exec():
            filename = dialog.selectedFiles()
            print(filename)
            if filename:
                csv_file_browser.setText(filename[0])
    elif button_id == 'save_browse':
        dialog.setDirectory(r'C:\\Desktop')
        dialog.setFileMode(QFileDialog.FileMode.Directory)
        dialog.setViewMode(QFileDialog.ViewMode.List)
        if dialog.exec():
            filename = dialog.selectedFiles()
            print(filename)
            if filename:
                save_file_browser.setText(filename[0])      
            
def load_data_from_csv(csv_filepath):
    headings, rows = [], []
    try:    
        with open(csv_filepath, encoding="utf8") as csv_file:
            for row in csv.reader(csv_file,delimiter=","):
                if row[0].lower() == "cumulative":
                    n=0
                    for row in csv.reader(csv_file, delimiter=","):
                        if row[0].lower() != "bounces" and row[0].lower() != "not interested":
                            
                            if row[0].lower() == "email sent":
                                headings.append("Sent")
                                
                            elif row[0].lower() == "delivered emails":
                                headings.append("Delivered")
                                
                            elif row[0].lower() == "maybe later":
                                headings.append("Maybe")
                            else:
                                heading = row[0]
                                headings.append(heading)
                                
                            data = row[-1]
                            if n == 0:
                                week_commence = row[1]
                            rows.append(data)
                            n+=1
                    #data = [headings,rows]
                    df = p.DataFrame([rows],columns=headings)
                    print(data)
                else:
                    print('couldnt find correct rows')
            
    except FileNotFoundError as e:
        print(e)
    if headings or rows != []:
        return headings, rows, week_commence, df
    

    else:
        raise ValueError('Your CSV does not contain headings and/or rows')
    
    
def page_setup(canvas, doc):
    w,h=A4
    canvas.saveState()
    canvas.setFont('Times-Roman',24)
    canvas.drawCentredString(w/2.0, h-50*mm, title)
    canvas.setFont('Times-Roman',9)
    
    canvas.drawCentredString(w/2,16*mm, "%s" % pageinfo)
    canvas.drawCentredString(20*mm,16*mm,"%s" % date.today())
    
    canvas.setStrokeColorRGB(0,119/255,181/255)
    canvas.setLineWidth(2)
    outer_box = [(8*mm,8*mm,w-8*mm,8*mm),(8*mm,h-8*mm,w-8*mm,h-8*mm),(8*mm,h-8*mm,8*mm,8*mm),(w-8*mm,h-8*mm,w-8*mm,8*mm)]
    canvas.lines(outer_box)#outer box draw
    
    canvas.drawImage("required/better_logo.png",w-78*mm,h-30*mm,width=62.5*mm,height=15*mm,mask='auto')
    
    canvas.restoreState()

def create_pdf(data,week_commence,save_location):
    
    data['Open Rate'] = np.where(int(data['Views'][0])<1, int(data['Views'][0]),int(data['Views'][0])/int(data['Delivered'][0])*100)
    data['Open Rate'] = '{0:.2f}%'.format(data['Open Rate'][0])
    print(data)
    data = data[['Sent','Open Rate','Replies','Interested','Maybe']]    


    doc = SimpleDocTemplate(save_location+'.pdf', pagesize=A4)
    
    
    # container for the 'Flowable' objects
    elements = []
    
    
    styleSheet = getSampleStyleSheet()
    columns = int(len(data.columns))
    
    
    w,h=A4
    width = (w-80)/columns
    resplit_columns = data.columns.tolist()
    resplit_rows = data.loc[0, :].values.flatten().tolist()
    final_data = []
    final_data.append(resplit_columns)
    final_data.append(resplit_rows)
    
    
    t1=Table(final_data,colWidths=width,style=[('BOX',(0,0),(-1,-1),1,colors.black),
    ('VALIGN',(0,0),(columns,1),'MIDDLE'),
    ('ALIGN',(0,0),(columns,1),'CENTER'),
    ('ALIGN',(0,0),(columns,1),'CENTER'),
    ],cornerRadii=[8,8,8,8])
    
    t2=Table(week_commence,hAlign=TA_CENTER,colWidths=(w-80),style=[
    ('VALIGN',(0,0),(0,0),'MIDDLE'),
    ('ALIGN',(0,0),(0,0),'CENTER'),
    ('TEXTCOLOR',(0,0),(0,0),colors.white),
    ('BACKGROUND',(0,0),(0,0),colors.HexColor('#0077b5'))]
             ,cornerRadii=[8,8,8,8])
    
    
    elements.append(Spacer(w,40*mm))
    elements.append(t2)
    elements.append(Spacer(w,3*mm))
    elements.append(t1)
    # write the document to disk
    
    
    doc.build(elements,onFirstPage=page_setup)
    

def read_csv():
    headings, rows, week_commence, df = load_data_from_csv(csv_file_browser.text())
    year = str(datetime.date.today().year)
    date_to_parse = (week_commence+' '+year)
    new_date = parse(date_to_parse)
    date_list = list(new_date.timetuple())
    
    
    if date_list[2] == 1 or date_list[2]==21 or date_list[2]==31:
        dates_for_display =  new_date.strftime('%A %dst %B %Y')
        
    elif date_list[2] == 2 or date_list[2]==22:
        dates_for_display =  new_date.strftime('%A %dnd %B %Y')
        
    elif date_list[2] == 3 or date_list[2]==23:
        dates_for_display =  new_date.strftime('%A %drd %B %Y')
        
    else:
        dates_for_display =  new_date.strftime('%A %dth %B %Y')
        
    week_commencing = [[dates_for_display]]    
    return (df,week_commencing)

def button_click(button_id):
    if button_id == 'generate':
        ##makes sure all inputs have been filled before generating save location or pdf
        if csv_file_browser.text() != '' and save_file_browser.text() != '' and file_name.text() != '' and dropdown.currentIndex() != -1:
            save_location = os.path.join(save_file_browser.text(),file_name.text())
            
            create_pdf(data=read_csv()[0],week_commence=read_csv()[1],save_location=save_location)
        else:
            print('Missing input')
            #error_msg('Missing Inputs')
        
    elif button_id == 'csv_browse':
        file_explorer(button_id)
        
    elif button_id == 'save_browse':
        file_explorer(button_id)
    
    return(button_id)
    
    
app = QApplication([])
app.setStyle('Fusion')

##creates window
window = QWidget()
window.setWindowTitle('CSV to PDF v0.02')
window.setGeometry(100,100,280,80)

##creates layout
layout = QGridLayout()


##creates labels
csv_file_browser = QLineEdit()

save_file_browser = QLineEdit()

file_name = QLineEdit()

#sets texts and add widgets to layout with span

layout.addWidget(csv_file_browser,0,0,1,2)
csv_file_browser.setPlaceholderText('Select CSV')

save_file_browser.setPlaceholderText('Select Save Location')
layout.addWidget(save_file_browser,1,0,1,2)

file_name.setPlaceholderText('Save As')
layout.addWidget(file_name,3,0,1,3)

#adds dropdown menu
dropdown = QComboBox()
dropdown.setPlaceholderText('Select Report Type')
dropdown.insertItems(1,['Weekly','To Date'])
layout.addWidget(dropdown,2,0,1,3)


##creates pushbuttons and adds them to the layout
generate_button = QPushButton('Generate')
generate_button.clicked.connect(button_click(partial(button_click,'generate')))

csv_browse_button = QPushButton('Browse')
csv_browse_button.clicked.connect(button_click(partial(button_click,'csv_browse')))

save_browse_button = QPushButton('Browse')
save_browse_button.clicked.connect(button_click(partial(button_click,'save_browse')))

apply_button = QPushButton('Apply')
apply_button.clicked.connect(button_click(partial(button_click,'apply')))

layout.addWidget(generate_button,4,0,1,3)
layout.addWidget(csv_browse_button,0,2)
layout.addWidget(save_browse_button,1,2)
layout.addWidget(apply_button,2,2,1,1)


window.setLayout(layout)
window.show()
sys.exit(app.exec())
