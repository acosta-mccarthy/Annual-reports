#!/usr/bin/env python3

"""Create and email annual report for Jan 1st

Author: Nina Acosta
"""

import psycopg2
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date, timedelta
yesterday = date.today() - timedelta(1)
year = yesterday.strftime ("%Y") #Pulls the year for the subject line and the body of the email

#SQL Query Part I & II - Holdings and Additions
q1='''SELECT
LEFT(location_code, 3) AS "LOCATION",

--PART I - Total holdings by location and statistical code
COUNT(CASE when icode2 = 'a' then 1 end)AS "Adult Fiction 2.1",
COUNT(CASE when icode2 = 'b' then 1 end)AS "Adult Non-Fiction 2.2",
COUNT(CASE when icode2 = 'a' OR icode2 = 'b' then 1 end)AS "Total Adult 2.3",
COUNT(CASE when icode2 = 'c' then 1 end)AS "Juvenile Fiction 2.4",
COUNT(CASE when icode2 = 'd' then 1 end)AS "Juvenile Non-Fiction 2.5",
COUNT(CASE when icode2 = 'c' OR icode2 = 'd' then 1 end)AS "Total Juvenile Books 2.6",
COUNT(CASE when icode2 = 'a' OR icode2 = 'b' OR icode2 = 'c' OR icode2 = 'd' then 1 end)AS "Total Books 2.7",
COUNT(CASE when icode2 = 'f' then 1 end)AS "Microform",
COUNT(CASE when icode2 = 'h' then 1 end)AS "Adult Sound Recording",
COUNT(CASE when icode2 = 'i' then 1 end)AS "Adult Videorecording",
COUNT(CASE when icode2 = 'j' then 1 end)AS "Media",
COUNT(CASE when icode2 = 'l' then 1 end)AS "Adult Software",
COUNT(CASE when icode2 = 'm' then 1 end)AS "Equipment/Relia",
COUNT(CASE when icode2 = 'n' then 1 end)AS "Supressed Item",
COUNT(CASE when icode2 = 'q' then 1 end)AS "Juvenile Video",
COUNT(CASE when icode2 = 'r' then 1 end)AS "Juvenile Audio",
COUNT(CASE when icode2 = 's' then 1 end)AS "Juvenile Other Media",
COUNT(CASE when icode2 = 't' then 1 end)AS "Juvenile Software",
COUNT(CASE when icode2 = 'z' then 1 end)AS "Vertical File",
COUNT(CASE when icode2 = 'n' OR icode2 = 'z' then 1 end)AS "All Other Print 2.10",
COUNT(CASE when icode2 = 'a' OR icode2 = 'b' OR icode2 = 'c' OR icode2 = 'd' OR icode2 = 'n' OR icode2 = 'z' then 1 end)AS "Total Print 2.12",
COUNT(CASE when icode2 = 'l' OR icode2 = 't' then 1 end)AS "Total Other Electronic Materials 2.19",
COUNT(CASE when icode2 = 'h' OR icode2 = 'r' then 1 end)AS "Total Sound Recording 2.21",
COUNT(CASE when icode2 = 'i' OR icode2 = 'q' then 1 end)AS "Total Videorecording 2.22",
COUNT(CASE when icode2 = 'f' OR icode2 = 'j' OR icode2 = 'm' OR icode2 = 's' then 1 end)AS "All Other Materials 2.23",
COUNT(CASE when icode2 = 'h' OR icode2 = 'r' OR icode2 = 'i' OR icode2 = 'q' OR icode2 = 'f' OR icode2 = 'j' OR icode2 = 'm' OR icode2 = 's' then 1 end)AS "Total Other Materials",


--PART II - Total additions (all holdings added during previous year) by location and statistical code
COUNT(CASE when icode2 = 'a' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Fiction Added",
COUNT(CASE when icode2 = 'b' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Non-Fiction Added",
COUNT(CASE when icode2 = 'c' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Fiction Added",
COUNT(CASE when icode2 = 'd' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Non-Fiction Added",
COUNT(CASE when (icode2 = 'a' OR icode2 = 'b' OR icode2 = 'c' OR icode2 = 'd') AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Cataloged Books added 2.27",
COUNT(CASE when icode2 = 'l' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Software Added",
COUNT(CASE when icode2 = 't' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Software Added",
COUNT(CASE when icode2 = 'f' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Microfilm Added",
COUNT(CASE when icode2 = 'h' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Sound Recording Added",
COUNT(CASE when icode2 = 'i' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Videorecording Added",
COUNT(CASE when icode2 = 'j' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Other Media Added",
COUNT(CASE when icode2 = 'm' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Equipment/Realia Added",
COUNT(CASE when icode2 = 'n' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Suppressed Items Added",
COUNT(CASE when icode2 = 'q' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Videorecording Added",
COUNT(CASE when icode2 = 'r' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Sound Recording Added",
COUNT(CASE when icode2 = 's' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Other Media Added",
COUNT(CASE when icode2 = 'z' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Vertical File Added",
COUNT(CASE when (icode2 = 'n' OR icode2 = 'z') AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "All Other Print Materials Added 2.28"

FROM
sierra_view.item_view

WHERE
location_code != '' AND
location_code NOT LIKE 'non%' AND
location_code NOT LIKE 'zzz%'
--Excludes any bad location codes

GROUP BY "LOCATION"
ORDER BY "LOCATION"
'''


#SQL Query Part III
q2='''SELECT

--PART III - Total borrowers by location and residency
LEFT(home_library_code, 3) AS "LOCATION",
COUNT(CASE when home_library_code IS NOT NULL AND ptype_code != '3' then 1 end)AS "Resident Borrowers 3.2",
COUNT(CASE when home_library_code IS NOT NULL AND ptype_code = '3' then 1 end)AS "Non-Resident Borrowers 3.3",
COUNT(CASE when home_library_code IS NOT NULL then 1 end)AS  "Total Number of Borrowers"

FROM
sierra_view.patron_view

WHERE
home_library_code != '' AND
home_library_code NOT LIKE 'non%' AND
home_library_code NOT LIKE 'zzz%'
--Excludes any bad location codes

GROUP BY "LOCATION"
ORDER BY "LOCATION"
'''

#Name of Excel File
excelfile = "C:/Users/staff/Desktop/SampleReportForMAIUG_"+ str(year)+".xlsx" #Adds the year to the end of the filename

# These are variables for the email that will be sent.
# This code uses placeholders, please add your own email server info
emailhost = 'email.server.midhudson.org'
emailuser = 'emailaddress@midhudson.org'
emailpass = '*******'
emailport = '587'
emailsubject = 'Annual Report ' + str(year)
emailmessage = '''***This is an automated email***


The ''' + str(year) + ''' annual report masterlist is attached.
This spreadsheet contains the Holdings, Additions, and Borrowers for all non-school district libraries.'''
emailfrom= 'emailaddress@midhudson.org'
emailto = 'nacosta@midhudson.org'


#This code uses placeholder info to connect to Sierra SQL server, please replace with your own info
conn = psycopg2.connect("dbname='iii' user='*****' host='000.000.000.000' port='1032' password='*****' sslmode='require'")

#Open session and runs both queries
cursor = conn.cursor()
cursor.execute(q1)
rows = cursor.fetchall()
cursor.execute(q2)
lines = cursor.fetchall()
conn.close()

#Create Excel file
import xlsxwriter
workbook = xlsxwriter.Workbook("C:/Users/staff/Desktop/SampleReportForMAIUG_"+ str(year)+".xlsx") #Filepath of where to save the spreadsheet
worksheet = workbook.add_worksheet()
worksheet.freeze_panes(1, 1)  # Freeze first row and first column.

#Formatting our Excel worksheet
worksheet.set_landscape()
worksheet.hide_gridlines(0)

#Formatting Cells
eformatheading= workbook.add_format({'text_wrap': False, 'valign': 'top', 'bold': True})
eformat= workbook.add_format({'text_wrap': True, 'valign': 'top'})
eformatlabel= workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'bold': True})
bold = workbook.add_format({'valign': 'top','bold': True})


# Setting the column widths
worksheet.set_column(0,0,8.00) #Library
worksheet.set_column(1,1,6.29) #Adult Fic
worksheet.set_column(2,2,9.71) #Adult Non-Fic
worksheet.set_column(3,3,8.14) #Total Adult
worksheet.set_column(4,4,7.71) #Juv Fic
worksheet.set_column(5,5,10.71) #Juv Non-fic
worksheet.set_column(6,6,8.57) #Total Juv
worksheet.set_column(7,7,8.57) #Total Books
worksheet.set_column(8,8,9.43) #Microform
worksheet.set_column(9,9,9.14) #Adult Sound
worksheet.set_column(10,10,13.71) #Adult Video
worksheet.set_column(11,11,6.00) #Media
worksheet.set_column(12,12,8.14) #Adult Software
worksheet.set_column(13,13,9.71) #Equip
worksheet.set_column(14,14,10.43) #Suppress
worksheet.set_column(15,15,7.57) #Juv Video
worksheet.set_column(16,16,7.43) #Juv Audio
worksheet.set_column(17,17,7.71) #Juv Other
worksheet.set_column(18,18,8.00) #Juv Software
worksheet.set_column(19,19,7.00) #Vert file
worksheet.set_column(20,20,8.29) #Other Print
worksheet.set_column(21,21,8.43) #Total Print
worksheet.set_column(22,22,6.14) #ebook
worksheet.set_column(23,23,13.00) #Audio Download
worksheet.set_column(24,24,13.86) #Total video Download
worksheet.set_column(25,25,12.57) #total other electronic
worksheet.set_column(26,26,10.57) #total sound
worksheet.set_column(27,27,14.00) #total Video
worksheet.set_column(28,28,8.71) #all Other
worksheet.set_column(29,29,8.43) #Total Other
worksheet.set_column(30,30,7.86) #Grand total
worksheet.set_column(31,31,6.71) #Adult Fic add
worksheet.set_column(32,32,9.53) #adult non-fic add
worksheet.set_column(33,33,7.43) #Juv fic add
worksheet.set_column(34,34,10.43) #Juv non-fic add
worksheet.set_column(35,35,9.71) #books add
worksheet.set_column(36,36,8.00) #adult software add
worksheet.set_column(37,37,8.00) #juv software add
worksheet.set_column(38,38,6.57) #ebooks add
worksheet.set_column(39,39,8.71) #electronic add
worksheet.set_column(40,40,9.00) #microfilm add
worksheet.set_column(41,41,10.86) #adult sound add
worksheet.set_column(42,42,13.86) #adult video add
worksheet.set_column(43,43,10.43) #Adult media add
worksheet.set_column(44,44,9.86) #equip add
worksheet.set_column(45,45,10.43) #suppress add
worksheet.set_column(46,46,13.71) #Juv video add
worksheet.set_column(47,47,13.43) #Juv sound add
worksheet.set_column(48,48,11.29) #Juv media add
worksheet.set_column(49,49,6.86) #vert file add
worksheet.set_column(50,50,5.86) #other media add
worksheet.set_column(51,51,12.86) #download audio add
worksheet.set_column(52,52,13.00) #other print add
worksheet.set_column(53,53,9.86) #other materials add
worksheet.set_column(54,54,6.00) #Total add
#worksheet.set_column(55,55,24.71) #library name
worksheet.set_column(55,55,9.71) #resident
worksheet.set_column(56,56,12.71) #non-resident
worksheet.set_column(57,57,10.00) #total borrowers


#Inserting a header

# Adding column labels
worksheet.write(0,0,"Library", eformatlabel)
worksheet.write(0,1,"Adult Fiction 2.1", eformatlabel)
worksheet.write(0,2,"Adult Non-Fiction 2.2", eformatlabel)
worksheet.write(0,3,"Total Adult 2.3", eformatlabel)
worksheet.write(0,4,"Juvenile Fiction 2.4", eformatlabel)
worksheet.write(0,5,"Juvenile Non-Fiction 2.5", eformatlabel)
worksheet.write(0,6,"Total Juvenile Books 2.6", eformatlabel)
worksheet.write(0,7,"Total Books 2.7", eformatlabel)
worksheet.write(0,8,"Microform", eformatlabel)
worksheet.write(0,9,"Adult Sound Recording", eformatlabel)
worksheet.write(0,10,"Adult Videorecording", eformatlabel)
worksheet.write(0,11,"Media", eformatlabel)
worksheet.write(0,12,"Adult Software", eformatlabel)
worksheet.write(0,13,"Equipment/Realia", eformatlabel)
worksheet.write(0,14,"Suppressed Item", eformatlabel)
worksheet.write(0,15,"Juvenile Video", eformatlabel)
worksheet.write(0,16,"Juvenile Audio", eformatlabel)
worksheet.write(0,17,"Juvenile Other Media", eformatlabel)
worksheet.write(0,18,"Juvenile Software", eformatlabel)
worksheet.write(0,19,"Vertical File", eformatlabel)
worksheet.write(0,20,"All Other Print 2.10", eformatlabel)
worksheet.write(0,21,"Total Print 2.12", eformatlabel)
worksheet.write(0,22,"eBook 2.13", eformatlabel)
worksheet.write(0,23,"Audio Downloadable Units 2.17", eformatlabel)
worksheet.write(0,24,"Total Videorecording Downloadable", eformatlabel)
worksheet.write(0,25,"Total Other Electronic Materials 2.19", eformatlabel)
worksheet.write(0,26,"Total Sound Recording 2.21", eformatlabel)
worksheet.write(0,27,"Total Videorecording 2.22", eformatlabel)
worksheet.write(0,28,"All Other Materials 2.23", eformatlabel)
worksheet.write(0,29,"Total Other Materials", eformatlabel)
worksheet.write(0,30,"Grand Total Holdings", eformatlabel)
worksheet.write(0,31,"Adult Fiction Added", eformatlabel)
worksheet.write(0,32,"Adult Non-Fiction Added", eformatlabel)
worksheet.write(0,33,"Juvenile Fiction Added", eformatlabel)
worksheet.write(0,34,"Juvenile Non-Fiction Added", eformatlabel)
worksheet.write(0,35,"Cataloged Books added 2.27", eformatlabel)
worksheet.write(0,36,"Adult Software Added", eformatlabel)
worksheet.write(0,37,"Juvenile Software Added", eformatlabel)
worksheet.write(0,38,"eBooks Added", eformatlabel)
worksheet.write(0,39,"Electronic Materials Added 2.29", eformatlabel)
worksheet.write(0,40,"Microfilm Added", eformatlabel)
worksheet.write(0,41,"Adult Sound Recording Added", eformatlabel)
worksheet.write(0,42,"Adult Videorecording Added", eformatlabel)
worksheet.write(0,43,"Adult Other Media Added", eformatlabel)
worksheet.write(0,44,"Equipment/Realia Added", eformatlabel)
worksheet.write(0,45,"Suppressed Items Added", eformatlabel)
worksheet.write(0,46,"Juvenile Videorecording Added", eformatlabel)
worksheet.write(0,47,"Juvenile Sound Recording Added", eformatlabel)
worksheet.write(0,48,"Juvenile Other Media Added", eformatlabel)
worksheet.write(0,49,"Vertical File Added", eformatlabel)
worksheet.write(0,50,"Other Media Added", eformatlabel)
worksheet.write(0,51,"Downloadable Audio Added", eformatlabel)
worksheet.write(0,52,"All Other Print Materials Added 2.28", eformatlabel)
worksheet.write(0,53,"All Other Materials Added 2.30", eformatlabel)
worksheet.write(0,54,"Total Added", eformatlabel)
#worksheet.write(0,55,"Library name", eformatlabel)
worksheet.write(0,55,"Resident Borrowers 3.2", eformatlabel)
worksheet.write(0,56,"Non-Resident Borrowers 3.3", eformatlabel)
worksheet.write(0,57,"Total Number of Borrowers", eformatlabel)

#Inserts comments into some cells
worksheet.write_comment('U1', 'Sum of O2 (Supressed Items) and T2 (Vertical File).')
worksheet.write_comment('V1', 'Sum of Total Print (Column H) and Total Other Print (Column U)')
worksheet.write_comment('W1', 'Supplied by Outreach/Database deparmtent')
worksheet.write_comment('X1', 'Supplied by Outreach/Database Department ')
worksheet.write_comment('Y1', 'As of Annual report 2013, this is zero.  If data is supplied will come from Outreach/Databases Department')


# Writing the report for staff to the Excel worksheet
for rownum, row in enumerate(rows):
    worksheet.write(rownum+1,0,row[0], eformatlabel)#Library
    worksheet.write(rownum+1,1,row[1],eformat)#AF
    worksheet.write(rownum+1,2,row[2],eformat)#ANF
    worksheet.write(rownum+1,3,row[3],eformat)#TotA
    worksheet.write(rownum+1,4,row[4],eformat)#JF
    worksheet.write(rownum+1,5,row[5],eformat)#JNF
    worksheet.write(rownum+1,6,row[6],eformat)#TotJ
    worksheet.write(rownum+1,7,row[7],eformat)#TotBook
    worksheet.write(rownum+1,8,row[8],eformat)#Micro
    worksheet.write(rownum+1,9,row[9],eformat)#ASR
    worksheet.write(rownum+1,10,row[10],eformat)#AVR
    worksheet.write(rownum+1,11,row[11],eformat)#AMed
    worksheet.write(rownum+1,12,row[12],eformat)#ASW
    worksheet.write(rownum+1,13,row[13],eformat)#Equip
    worksheet.write(rownum+1,14,row[14],eformat)#Supp
    worksheet.write(rownum+1,15,row[15],eformat)#JVR
    worksheet.write(rownum+1,16,row[16],eformat)#JSR
    worksheet.write(rownum+1,17,row[17],eformat)#JotherMed
    worksheet.write(rownum+1,18,row[18],eformat)#JSW
    worksheet.write(rownum+1,19,row[19],eformat)#VF
    worksheet.write(rownum+1,20,row[20],eformat)#AllOtherPrint
    worksheet.write(rownum+1,21,row[21],eformat)#TotPrint
    #blank column for ebooks[22]
    #blank column for downloadable audio[23]
    #blank column for downloadable video[24]
    worksheet.write(rownum+1,25,row[22],eformat)#TotOtherEmat
    worksheet.write(rownum+1,26,row[23],eformat)#TotSR
    worksheet.write(rownum+1,27,row[24],eformat)#TotVR
    worksheet.write(rownum+1,28,row[25],eformat)#AllOtherMats
    worksheet.write(rownum+1,29,row[26],eformat)#TotOtherMats
    #Formula for Grand total holdings column [30]~OK~
    worksheet.write_formula('AE2','=SUM(V2,W2, Z2, AD2)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE3','=SUM(V3,W3, Z3, AD3)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE4','=SUM(V4,W4, Z4, AD4)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE5','=SUM(V5,W5, Z5, AD5)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE6','=SUM(V6,W6, Z6, AD6)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE7','=SUM(V7,W7, Z7, AD7)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE8','=SUM(V8,W8, Z8, AD8)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE9','=SUM(V9,W9, Z9, AD9)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE10','=SUM(V10,W10, Z10, AD10)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE11','=SUM(V11,W11, Z11, AD11)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE12','=SUM(V12,W12, Z12, AD12)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE13','=SUM(V13,W13, Z13, AD13)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE14','=SUM(V14,W14, Z14, AD14)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE15','=SUM(V15,W15, Z15, AD15)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE16','=SUM(V16,W16, Z16, AD16)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE17','=SUM(V17,W17, Z17, AD17)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE18','=SUM(V18,W18, Z18, AD18)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE19','=SUM(V19,W19, Z19, AD19)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE20','=SUM(V20,W20, Z20, AD20)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE21','=SUM(V21,W21, Z21, AD21)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE22','=SUM(V22,W22, Z22, AD22)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE23','=SUM(V23,W23, Z23, AD23)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE24','=SUM(V24,W24, Z24, AD24)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE25','=SUM(V25,W25, Z25, AD25)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE26','=SUM(V26,W26, Z26, AD26)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE27','=SUM(V27,W27, Z27, AD27)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE28','=SUM(V28,W28, Z28, AD28)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE29','=SUM(V29,W29, Z29, AD29)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE30','=SUM(V30,W30, Z30, AD30)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE31','=SUM(V31,W31, Z31, AD31)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE32','=SUM(V32,W32, Z32, AD32)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE33','=SUM(V33,W33, Z33, AD33)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE34','=SUM(V34,W34, Z34, AD34)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE35','=SUM(V35,W35, Z35, AD35)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE36','=SUM(V36,W36, Z36, AD36)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE37','=SUM(V37,W37, Z37, AD37)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE38','=SUM(V38,W38, Z38, AD38)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE39','=SUM(V39,W39, Z39, AD39)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE40','=SUM(V40,W40, Z40, AD40)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE41','=SUM(V41,W41, Z41, AD41)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE42','=SUM(V42,W42, Z42, AD42)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE43','=SUM(V43,W43, Z43, AD43)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE44','=SUM(V44,W44, Z44, AD44)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE45','=SUM(V45,W45, Z45, AD45)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE46','=SUM(V46,W46, Z46, AD46)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE47','=SUM(V47,W47, Z47, AD47)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE48','=SUM(V48,W48, Z48, AD48)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE49','=SUM(V49,W49, Z49, AD49)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE50','=SUM(V50,W50, Z50, AD50)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE51','=SUM(V51,W51, Z51, AD51)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE52','=SUM(V52,W52, Z52, AD52)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE53','=SUM(V53,W53, Z53, AD53)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE54','=SUM(V54,W54, Z54, AD54)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE55','=SUM(V55,W55, Z55, AD55)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE56','=SUM(V56,W56, Z56, AD56)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE57','=SUM(V57,W57, Z57, AD57)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE58','=SUM(V58,W58, Z58, AD58)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE59','=SUM(V59,W59, Z59, AD59)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE60','=SUM(V60,W60, Z60, AD60)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE61','=SUM(V61,W61, Z61, AD61)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE62','=SUM(V62,W62, Z62, AD62)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE63','=SUM(V63,W63, Z63, AD63)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE64','=SUM(V64,W64, Z64, AD64)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE65','=SUM(V65,W65, Z65, AD65)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE66','=SUM(V66,W66, Z66, AD66)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE67','=SUM(V67,W67, Z67, AD67)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE68','=SUM(V68,W68, Z68, AD68)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE69','=SUM(V69,W69, Z69, AD69)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE70','=SUM(V70,W70, Z70, AD70)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE71','=SUM(V71,W71, Z71, AD71)', eformat)#Grand Total Holdings
    worksheet.write_formula('AE72','=SUM(V72,W72, Z72, AD72)', eformat)#Grand Total Holdings

    worksheet.write(rownum+1,31,row[27],eformat)#AF+
    worksheet.write(rownum+1,32,row[28],eformat)#ANF+
    worksheet.write(rownum+1,33,row[29],eformat)#JF+
    worksheet.write(rownum+1,34,row[30],eformat)#JNF+
    worksheet.write(rownum+1,35,row[31],eformat)#Book+
    worksheet.write(rownum+1,36,row[32],eformat)#ASW+
    worksheet.write(rownum+1,37,row[33],eformat)#JSW+
    #blank column for ebooks added [38]
    #Formula for electronic materials added column [39]~OK~
    worksheet.write_formula('AN2','=SUM(AK2,AL2,AM2)', eformat)#Electronic materials added
    worksheet.write_formula('AN3','=SUM(AK3,AL3,AM3)', eformat)#Electronic materials added
    worksheet.write_formula('AN4','=SUM(AK4,AL4,AM4)', eformat)#Electronic materials added
    worksheet.write_formula('AN5','=SUM(AK5,AL5,AM5)', eformat)#Electronic materials added
    worksheet.write_formula('AN6','=SUM(AK6,AL6,AM6)', eformat)#Electronic materials added
    worksheet.write_formula('AN7','=SUM(AK7,AL7,AM7)', eformat)#Electronic materials added
    worksheet.write_formula('AN8','=SUM(AK8,AL8,AM8)', eformat)#Electronic materials added
    worksheet.write_formula('AN9','=SUM(AK9,AL9,AM9)', eformat)#Electronic materials added
    worksheet.write_formula('AN10','=SUM(AK10,AL10,AM10)', eformat)#Electronic materials added
    worksheet.write_formula('AN11','=SUM(AK11,AL11,AM11)', eformat)#Electronic materials added
    worksheet.write_formula('AN12','=SUM(AK12,AL12,AM12)', eformat)#Electronic materials added
    worksheet.write_formula('AN13','=SUM(AK13,AL13,AM13)', eformat)#Electronic materials added
    worksheet.write_formula('AN14','=SUM(AK14,AL14,AM14)', eformat)#Electronic materials added
    worksheet.write_formula('AN15','=SUM(AK15,AL15,AM15)', eformat)#Electronic materials added
    worksheet.write_formula('AN16','=SUM(AK16,AL16,AM16)', eformat)#Electronic materials added
    worksheet.write_formula('AN17','=SUM(AK17,AL17,AM17)', eformat)#Electronic materials added
    worksheet.write_formula('AN18','=SUM(AK18,AL18,AM18)', eformat)#Electronic materials added
    worksheet.write_formula('AN19','=SUM(AK19,AL19,AM19)', eformat)#Electronic materials added
    worksheet.write_formula('AN20','=SUM(AK20,AL20,AM20)', eformat)#Electronic materials added
    worksheet.write_formula('AN21','=SUM(AK21,AL21,AM21)', eformat)#Electronic materials added
    worksheet.write_formula('AN22','=SUM(AK22,AL22,AM22)', eformat)#Electronic materials added
    worksheet.write_formula('AN23','=SUM(AK23,AL23,AM23)', eformat)#Electronic materials added
    worksheet.write_formula('AN24','=SUM(AK24,AL24,AM24)', eformat)#Electronic materials added
    worksheet.write_formula('AN25','=SUM(AK25,AL25,AM25)', eformat)#Electronic materials added
    worksheet.write_formula('AN26','=SUM(AK26,AL26,AM26)', eformat)#Electronic materials added
    worksheet.write_formula('AN27','=SUM(AK27,AL27,AM27)', eformat)#Electronic materials added
    worksheet.write_formula('AN28','=SUM(AK28,AL28,AM28)', eformat)#Electronic materials added
    worksheet.write_formula('AN29','=SUM(AK29,AL29,AM29)', eformat)#Electronic materials added
    worksheet.write_formula('AN30','=SUM(AK30,AL30,AM30)', eformat)#Electronic materials added
    worksheet.write_formula('AN31','=SUM(AK31,AL31,AM31)', eformat)#Electronic materials added
    worksheet.write_formula('AN32','=SUM(AK32,AL32,AM32)', eformat)#Electronic materials added
    worksheet.write_formula('AN33','=SUM(AK33,AL33,AM33)', eformat)#Electronic materials added
    worksheet.write_formula('AN34','=SUM(AK34,AL34,AM34)', eformat)#Electronic materials added
    worksheet.write_formula('AN35','=SUM(AK35,AL35,AM35)', eformat)#Electronic materials added
    worksheet.write_formula('AN36','=SUM(AK36,AL36,AM36)', eformat)#Electronic materials added
    worksheet.write_formula('AN37','=SUM(AK37,AL37,AM37)', eformat)#Electronic materials added
    worksheet.write_formula('AN38','=SUM(AK38,AL38,AM38)', eformat)#Electronic materials added
    worksheet.write_formula('AN39','=SUM(AK39,AL39,AM39)', eformat)#Electronic materials added
    worksheet.write_formula('AN40','=SUM(AK40,AL40,AM40)', eformat)#Electronic materials added
    worksheet.write_formula('AN41','=SUM(AK41,AL41,AM41)', eformat)#Electronic materials added
    worksheet.write_formula('AN42','=SUM(AK42,AL42,AM42)', eformat)#Electronic materials added
    worksheet.write_formula('AN43','=SUM(AK43,AL43,AM43)', eformat)#Electronic materials added
    worksheet.write_formula('AN44','=SUM(AK44,AL44,AM44)', eformat)#Electronic materials added
    worksheet.write_formula('AN45','=SUM(AK45,AL45,AM45)', eformat)#Electronic materials added
    worksheet.write_formula('AN46','=SUM(AK46,AL46,AM46)', eformat)#Electronic materials added
    worksheet.write_formula('AN47','=SUM(AK47,AL47,AM47)', eformat)#Electronic materials added
    worksheet.write_formula('AN48','=SUM(AK48,AL48,AM48)', eformat)#Electronic materials added
    worksheet.write_formula('AN49','=SUM(AK49,AL49,AM49)', eformat)#Electronic materials added
    worksheet.write_formula('AN50','=SUM(AK50,AL50,AM50)', eformat)#Electronic materials added
    worksheet.write_formula('AN51','=SUM(AK51,AL51,AM51)', eformat)#Electronic materials added
    worksheet.write_formula('AN52','=SUM(AK52,AL52,AM52)', eformat)#Electronic materials added
    worksheet.write_formula('AN53','=SUM(AK53,AL53,AM53)', eformat)#Electronic materials added
    worksheet.write_formula('AN54','=SUM(AK54,AL54,AM54)', eformat)#Electronic materials added
    worksheet.write_formula('AN55','=SUM(AK55,AL55,AM55)', eformat)#Electronic materials added
    worksheet.write_formula('AN56','=SUM(AK56,AL56,AM56)', eformat)#Electronic materials added
    worksheet.write_formula('AN57','=SUM(AK57,AL57,AM57)', eformat)#Electronic materials added
    worksheet.write_formula('AN58','=SUM(AK58,AL58,AM58)', eformat)#Electronic materials added
    worksheet.write_formula('AN59','=SUM(AK59,AL59,AM59)', eformat)#Electronic materials added
    worksheet.write_formula('AN60','=SUM(AK60,AL60,AM60)', eformat)#Electronic materials added
    worksheet.write_formula('AN61','=SUM(AK61,AL61,AM61)', eformat)#Electronic materials added
    worksheet.write_formula('AN62','=SUM(AK62,AL62,AM62)', eformat)#Electronic materials added
    worksheet.write_formula('AN63','=SUM(AK63,AL63,AM63)', eformat)#Electronic materials added
    worksheet.write_formula('AN64','=SUM(AK64,AL64,AM64)', eformat)#Electronic materials added
    worksheet.write_formula('AN65','=SUM(AK65,AL65,AM65)', eformat)#Electronic materials added
    worksheet.write_formula('AN66','=SUM(AK66,AL66,AM66)', eformat)#Electronic materials added
    worksheet.write_formula('AN67','=SUM(AK67,AL67,AM67)', eformat)#Electronic materials added
    worksheet.write_formula('AN68','=SUM(AK68,AL68,AM68)', eformat)#Electronic materials added
    worksheet.write_formula('AN69','=SUM(AK69,AL69,AM69)', eformat)#Electronic materials added
    worksheet.write_formula('AN70','=SUM(AK70,AL70,AM70)', eformat)#Electronic materials added
    worksheet.write_formula('AN71','=SUM(AK71,AL71,AM71)', eformat)#Electronic materials added
    worksheet.write_formula('AN72','=SUM(AK72,AL72,AM72)', eformat)#Electronic materials added

    worksheet.write(rownum+1,40,row[34],eformat)#Micro+
    worksheet.write(rownum+1,41,row[35],eformat)#ASR+
    worksheet.write(rownum+1,42,row[36],eformat)#AVR+
    worksheet.write(rownum+1,43,row[37],eformat)#AOtherMed+
    worksheet.write(rownum+1,44,row[38],eformat)#Equip+
    worksheet.write(rownum+1,45,row[39],eformat)#Supp+
    worksheet.write(rownum+1,46,row[40],eformat)#JVR+
    worksheet.write(rownum+1,47,row[41],eformat)#JSR+
    worksheet.write(rownum+1,48,row[42],eformat)#JotherMed+
    worksheet.write(rownum+1,49,row[43],eformat)#VF+
    #blank column for other media added [50]
    #blank column for Downloadable audio added [51]
    worksheet.write(rownum+1,52,row[44],eformat)#OtherPrint+

    #Formula for All other materials added column [53]~OK~
    worksheet.write_formula('BB2', '=SUM(AO2:AS2,AU2:AW2,AY2:AZ2)', eformat)#All other materials added
    worksheet.write_formula('BB3', '=SUM(AO3:AS3,AU3:AW3,AY3:AZ3)', eformat)#All other materials added
    worksheet.write_formula('BB4', '=SUM(AO4:AS4,AU4:AW4,AY4:AZ4)', eformat)#All other materials added
    worksheet.write_formula('BB5', '=SUM(AO5:AS5,AU5:AW5,AY5:AZ5)', eformat)#All other materials added
    worksheet.write_formula('BB6', '=SUM(AO6:AS6,AU6:AW6,AY6:AZ6)', eformat)#All other materials added
    worksheet.write_formula('BB7', '=SUM(AO7:AS7,AU7:AW7,AY7:AZ7)', eformat)#All other materials added
    worksheet.write_formula('BB8', '=SUM(AO8:AS8,AU8:AW8,AY8:AZ8)', eformat)#All other materials added
    worksheet.write_formula('BB9', '=SUM(AO9:AS9,AU9:AW9,AY9:AZ9)', eformat)#All other materials added
    worksheet.write_formula('BB10', '=SUM(AO10:AS10,AU10:AW10,AY10:AZ10)', eformat)#All other materials added
    worksheet.write_formula('BB11', '=SUM(AO11:AS11,AU11:AW11,AY11:AZ11)', eformat)#All other materials added
    worksheet.write_formula('BB12', '=SUM(AO12:AS12,AU12:AW12,AY12:AZ12)', eformat)#All other materials added
    worksheet.write_formula('BB13', '=SUM(AO13:AS13,AU13:AW13,AY13:AZ13)', eformat)#All other materials added
    worksheet.write_formula('BB14', '=SUM(AO14:AS14,AU14:AW14,AY14:AZ14)', eformat)#All other materials added
    worksheet.write_formula('BB15', '=SUM(AO15:AS15,AU15:AW15,AY15:AZ15)', eformat)#All other materials added
    worksheet.write_formula('BB16', '=SUM(AO16:AS16,AU16:AW16,AY16:AZ16)', eformat)#All other materials added
    worksheet.write_formula('BB17', '=SUM(AO17:AS17,AU17:AW17,AY17:AZ17)', eformat)#All other materials added
    worksheet.write_formula('BB18', '=SUM(AO18:AS18,AU18:AW18,AY18:AZ18)', eformat)#All other materials added
    worksheet.write_formula('BB19', '=SUM(AO19:AS19,AU19:AW19,AY19:AZ19)', eformat)#All other materials added
    worksheet.write_formula('BB20', '=SUM(AO20:AS20,AU20:AW20,AY20:AZ20)', eformat)#All other materials added
    worksheet.write_formula('BB21', '=SUM(AO21:AS21,AU21:AW21,AY21:AZ21)', eformat)#All other materials added
    worksheet.write_formula('BB22', '=SUM(AO22:AS22,AU22:AW22,AY22:AZ22)', eformat)#All other materials added
    worksheet.write_formula('BB23', '=SUM(AO23:AS23,AU23:AW23,AY23:AZ23)', eformat)#All other materials added
    worksheet.write_formula('BB24', '=SUM(AO24:AS24,AU24:AW24,AY24:AZ24)', eformat)#All other materials added
    worksheet.write_formula('BB25', '=SUM(AO25:AS25,AU25:AW25,AY25:AZ25)', eformat)#All other materials added
    worksheet.write_formula('BB26', '=SUM(AO26:AS26,AU26:AW26,AY26:AZ26)', eformat)#All other materials added
    worksheet.write_formula('BB27', '=SUM(AO27:AS27,AU27:AW27,AY27:AZ27)', eformat)#All other materials added
    worksheet.write_formula('BB28', '=SUM(AO28:AS28,AU28:AW28,AY28:AZ28)', eformat)#All other materials added
    worksheet.write_formula('BB29', '=SUM(AO29:AS29,AU29:AW29,AY29:AZ29)', eformat)#All other materials added
    worksheet.write_formula('BB30', '=SUM(AO30:AS30,AU30:AW30,AY30:AZ30)', eformat)#All other materials added
    worksheet.write_formula('BB31', '=SUM(AO31:AS31,AU31:AW31,AY31:AZ31)', eformat)#All other materials added
    worksheet.write_formula('BB32', '=SUM(AO32:AS32,AU32:AW32,AY32:AZ32)', eformat)#All other materials added
    worksheet.write_formula('BB33', '=SUM(AO33:AS33,AU33:AW33,AY33:AZ33)', eformat)#All other materials added
    worksheet.write_formula('BB34', '=SUM(AO34:AS34,AU34:AW34,AY34:AZ34)', eformat)#All other materials added
    worksheet.write_formula('BB35', '=SUM(AO35:AS35,AU35:AW35,AY35:AZ35)', eformat)#All other materials added
    worksheet.write_formula('BB36', '=SUM(AO36:AS36,AU36:AW36,AY36:AZ36)', eformat)#All other materials added
    worksheet.write_formula('BB37', '=SUM(AO37:AS37,AU37:AW37,AY37:AZ37)', eformat)#All other materials added
    worksheet.write_formula('BB38', '=SUM(AO38:AS38,AU38:AW38,AY38:AZ38)', eformat)#All other materials added
    worksheet.write_formula('BB39', '=SUM(AO39:AS39,AU39:AW39,AY39:AZ39)', eformat)#All other materials added
    worksheet.write_formula('BB40', '=SUM(AO40:AS40,AU40:AW40,AY40:AZ40)', eformat)#All other materials added
    worksheet.write_formula('BB41', '=SUM(AO41:AS41,AU41:AW41,AY41:AZ41)', eformat)#All other materials added
    worksheet.write_formula('BB42', '=SUM(AO42:AS42,AU42:AW42,AY42:AZ42)', eformat)#All other materials added
    worksheet.write_formula('BB43', '=SUM(AO43:AS43,AU43:AW43,AY43:AZ43)', eformat)#All other materials added
    worksheet.write_formula('BB44', '=SUM(AO44:AS44,AU44:AW44,AY44:AZ44)', eformat)#All other materials added
    worksheet.write_formula('BB45', '=SUM(AO45:AS45,AU45:AW45,AY45:AZ45)', eformat)#All other materials added
    worksheet.write_formula('BB46', '=SUM(AO46:AS46,AU46:AW46,AY46:AZ46)', eformat)#All other materials added
    worksheet.write_formula('BB47', '=SUM(AO47:AS47,AU47:AW47,AY47:AZ47)', eformat)#All other materials added
    worksheet.write_formula('BB48', '=SUM(AO48:AS48,AU48:AW48,AY48:AZ48)', eformat)#All other materials added
    worksheet.write_formula('BB49', '=SUM(AO49:AS49,AU49:AW49,AY49:AZ49)', eformat)#All other materials added
    worksheet.write_formula('BB50', '=SUM(AO50:AS50,AU50:AW50,AY50:AZ50)', eformat)#All other materials added
    worksheet.write_formula('BB51', '=SUM(AO51:AS51,AU51:AW51,AY51:AZ51)', eformat)#All other materials added
    worksheet.write_formula('BB52', '=SUM(AO52:AS52,AU52:AW52,AY52:AZ52)', eformat)#All other materials added
    worksheet.write_formula('BB53', '=SUM(AO53:AS53,AU53:AW53,AY53:AZ53)', eformat)#All other materials added
    worksheet.write_formula('BB54', '=SUM(AO54:AS54,AU54:AW54,AY54:AZ54)', eformat)#All other materials added
    worksheet.write_formula('BB55', '=SUM(AO55:AS55,AU55:AW55,AY55:AZ55)', eformat)#All other materials added
    worksheet.write_formula('BB56', '=SUM(AO56:AS56,AU56:AW56,AY56:AZ56)', eformat)#All other materials added
    worksheet.write_formula('BB57', '=SUM(AO57:AS57,AU57:AW57,AY57:AZ57)', eformat)#All other materials added
    worksheet.write_formula('BB58', '=SUM(AO58:AS58,AU58:AW58,AY58:AZ58)', eformat)#All other materials added
    worksheet.write_formula('BB59', '=SUM(AO59:AS59,AU59:AW59,AY59:AZ59)', eformat)#All other materials added
    worksheet.write_formula('BB60', '=SUM(AO60:AS60,AU60:AW60,AY60:AZ60)', eformat)#All other materials added
    worksheet.write_formula('BB61', '=SUM(AO61:AS61,AU61:AW61,AY61:AZ61)', eformat)#All other materials added
    worksheet.write_formula('BB62', '=SUM(AO62:AS62,AU62:AW62,AY62:AZ62)', eformat)#All other materials added
    worksheet.write_formula('BB63', '=SUM(AO63:AS63,AU63:AW63,AY63:AZ63)', eformat)#All other materials added
    worksheet.write_formula('BB64', '=SUM(AO64:AS64,AU64:AW64,AY64:AZ64)', eformat)#All other materials added
    worksheet.write_formula('BB65', '=SUM(AO65:AS65,AU65:AW65,AY65:AZ65)', eformat)#All other materials added
    worksheet.write_formula('BB66', '=SUM(AO66:AS66,AU66:AW66,AY66:AZ66)', eformat)#All other materials added
    worksheet.write_formula('BB67', '=SUM(AO67:AS67,AU67:AW67,AY67:AZ67)', eformat)#All other materials added
    worksheet.write_formula('BB68', '=SUM(AO68:AS68,AU68:AW68,AY68:AZ68)', eformat)#All other materials added
    worksheet.write_formula('BB69', '=SUM(AO69:AS69,AU69:AW69,AY69:AZ69)', eformat)#All other materials added
    worksheet.write_formula('BB70', '=SUM(AO70:AS70,AU70:AW70,AY70:AZ70)', eformat)#All other materials added
    worksheet.write_formula('BB71', '=SUM(AO71:AS71,AU71:AW71,AY71:AZ71)', eformat)#All other materials added
    worksheet.write_formula('BB72', '=SUM(AO72:AS72,AU72:AW72,AY72:AZ72)', eformat)#All other materials added


    #Formula for Total Added Column [54]~OK~
    worksheet.write_formula('BC2', '=SUM(AJ2,AN2,BA2,BB2 )', eformat)#Total added
    worksheet.write_formula('BC3', '=SUM(AJ3,AN3,BA3,BB3 )', eformat)#Total added
    worksheet.write_formula('BC4', '=SUM(AJ4,AN4,BA4,BB4 )', eformat)#Total added
    worksheet.write_formula('BC5', '=SUM(AJ5,AN5,BA5,BB5 )', eformat)#Total added
    worksheet.write_formula('BC6', '=SUM(AJ6,AN6,BA6,BB6 )', eformat)#Total added
    worksheet.write_formula('BC7', '=SUM(AJ7,AN7,BA7,BB7 )', eformat)#Total added
    worksheet.write_formula('BC8', '=SUM(AJ8,AN8,BA8,BB8 )', eformat)#Total added
    worksheet.write_formula('BC9', '=SUM(AJ9,AN9,BA9,BB9 )', eformat)#Total added
    worksheet.write_formula('BC10', '=SUM(AJ10,AN10,BA10,BB10 )', eformat)#Total added
    worksheet.write_formula('BC11', '=SUM(AJ11,AN11,BA11,BB11 )', eformat)#Total added
    worksheet.write_formula('BC12', '=SUM(AJ12,AN12,BA12,BB12 )', eformat)#Total added
    worksheet.write_formula('BC13', '=SUM(AJ13,AN13,BA13,BB13 )', eformat)#Total added
    worksheet.write_formula('BC14', '=SUM(AJ14,AN14,BA14,BB14 )', eformat)#Total added
    worksheet.write_formula('BC15', '=SUM(AJ15,AN15,BA15,BB15 )', eformat)#Total added
    worksheet.write_formula('BC16', '=SUM(AJ16,AN16,BA16,BB16 )', eformat)#Total added
    worksheet.write_formula('BC17', '=SUM(AJ17,AN17,BA17,BB17 )', eformat)#Total added
    worksheet.write_formula('BC18', '=SUM(AJ18,AN18,BA18,BB18 )', eformat)#Total added
    worksheet.write_formula('BC19', '=SUM(AJ19,AN19,BA19,BB19 )', eformat)#Total added
    worksheet.write_formula('BC20', '=SUM(AJ20,AN20,BA20,BB20 )', eformat)#Total added
    worksheet.write_formula('BC21', '=SUM(AJ21,AN21,BA21,BB21 )', eformat)#Total added
    worksheet.write_formula('BC22', '=SUM(AJ22,AN22,BA22,BB22 )', eformat)#Total added
    worksheet.write_formula('BC23', '=SUM(AJ23,AN23,BA23,BB23 )', eformat)#Total added
    worksheet.write_formula('BC24', '=SUM(AJ24,AN24,BA24,BB24 )', eformat)#Total added
    worksheet.write_formula('BC25', '=SUM(AJ25,AN25,BA25,BB25 )', eformat)#Total added
    worksheet.write_formula('BC26', '=SUM(AJ26,AN26,BA26,BB26 )', eformat)#Total added
    worksheet.write_formula('BC27', '=SUM(AJ27,AN27,BA27,BB27 )', eformat)#Total added
    worksheet.write_formula('BC28', '=SUM(AJ28,AN28,BA28,BB28 )', eformat)#Total added
    worksheet.write_formula('BC29', '=SUM(AJ29,AN29,BA29,BB29 )', eformat)#Total added
    worksheet.write_formula('BC30', '=SUM(AJ30,AN30,BA30,BB30 )', eformat)#Total added
    worksheet.write_formula('BC31', '=SUM(AJ31,AN31,BA31,BB31 )', eformat)#Total added
    worksheet.write_formula('BC32', '=SUM(AJ32,AN32,BA32,BB32 )', eformat)#Total added
    worksheet.write_formula('BC33', '=SUM(AJ33,AN33,BA33,BB33 )', eformat)#Total added
    worksheet.write_formula('BC34', '=SUM(AJ34,AN34,BA34,BB34 )', eformat)#Total added
    worksheet.write_formula('BC35', '=SUM(AJ35,AN35,BA35,BB35 )', eformat)#Total added
    worksheet.write_formula('BC36', '=SUM(AJ36,AN36,BA36,BB36 )', eformat)#Total added
    worksheet.write_formula('BC37', '=SUM(AJ37,AN37,BA37,BB37 )', eformat)#Total added
    worksheet.write_formula('BC38', '=SUM(AJ38,AN38,BA38,BB38 )', eformat)#Total added
    worksheet.write_formula('BC39', '=SUM(AJ39,AN39,BA39,BB39 )', eformat)#Total added
    worksheet.write_formula('BC40', '=SUM(AJ40,AN40,BA40,BB40 )', eformat)#Total added
    worksheet.write_formula('BC41', '=SUM(AJ41,AN41,BA41,BB41 )', eformat)#Total added
    worksheet.write_formula('BC42', '=SUM(AJ42,AN42,BA42,BB42 )', eformat)#Total added
    worksheet.write_formula('BC43', '=SUM(AJ43,AN43,BA43,BB43 )', eformat)#Total added
    worksheet.write_formula('BC44', '=SUM(AJ44,AN44,BA44,BB44 )', eformat)#Total added
    worksheet.write_formula('BC45', '=SUM(AJ45,AN45,BA45,BB45 )', eformat)#Total added
    worksheet.write_formula('BC46', '=SUM(AJ46,AN46,BA46,BB46 )', eformat)#Total added
    worksheet.write_formula('BC47', '=SUM(AJ47,AN47,BA47,BB47 )', eformat)#Total added
    worksheet.write_formula('BC48', '=SUM(AJ48,AN48,BA48,BB48 )', eformat)#Total added
    worksheet.write_formula('BC49', '=SUM(AJ49,AN49,BA49,BB49 )', eformat)#Total added
    worksheet.write_formula('BC50', '=SUM(AJ50,AN50,BA50,BB50 )', eformat)#Total added
    worksheet.write_formula('BC51', '=SUM(AJ51,AN51,BA51,BB51 )', eformat)#Total added
    worksheet.write_formula('BC52', '=SUM(AJ52,AN52,BA52,BB52 )', eformat)#Total added
    worksheet.write_formula('BC53', '=SUM(AJ53,AN53,BA53,BB53 )', eformat)#Total added
    worksheet.write_formula('BC54', '=SUM(AJ54,AN54,BA54,BB54 )', eformat)#Total added
    worksheet.write_formula('BC55', '=SUM(AJ55,AN55,BA55,BB55 )', eformat)#Total added
    worksheet.write_formula('BC56', '=SUM(AJ56,AN56,BA56,BB56 )', eformat)#Total added
    worksheet.write_formula('BC57', '=SUM(AJ57,AN57,BA57,BB57 )', eformat)#Total added
    worksheet.write_formula('BC58', '=SUM(AJ58,AN58,BA58,BB58 )', eformat)#Total added
    worksheet.write_formula('BC59', '=SUM(AJ59,AN59,BA59,BB59 )', eformat)#Total added
    worksheet.write_formula('BC60', '=SUM(AJ60,AN60,BA60,BB60 )', eformat)#Total added
    worksheet.write_formula('BC61', '=SUM(AJ61,AN61,BA61,BB61 )', eformat)#Total added
    worksheet.write_formula('BC62', '=SUM(AJ62,AN62,BA62,BB62 )', eformat)#Total added
    worksheet.write_formula('BC63', '=SUM(AJ63,AN63,BA63,BB63 )', eformat)#Total added
    worksheet.write_formula('BC64', '=SUM(AJ64,AN64,BA64,BB64 )', eformat)#Total added
    worksheet.write_formula('BC65', '=SUM(AJ65,AN65,BA65,BB65 )', eformat)#Total added
    worksheet.write_formula('BC66', '=SUM(AJ66,AN66,BA66,BB66 )', eformat)#Total added
    worksheet.write_formula('BC67', '=SUM(AJ67,AN67,BA67,BB67 )', eformat)#Total added
    worksheet.write_formula('BC68', '=SUM(AJ68,AN68,BA68,BB68 )', eformat)#Total added
    worksheet.write_formula('BC69', '=SUM(AJ69,AN69,BA69,BB69 )', eformat)#Total added
    worksheet.write_formula('BC70', '=SUM(AJ70,AN70,BA70,BB70 )', eformat)#Total added
    worksheet.write_formula('BC71', '=SUM(AJ71,AN71,BA71,BB71 )', eformat)#Total added
    worksheet.write_formula('BC72', '=SUM(AJ72,AN72,BA72,BB72 )', eformat)#Total added
for rownum, row in enumerate(lines): #pulls borrower data from separate SQL query
    #worksheet.write(rownum+1,55,row[0],eformat)#Libraries
    worksheet.write(rownum+1,55,row[1],eformat)#Resident
    worksheet.write(rownum+1,56,row[2],eformat)#Non-resident
    worksheet.write(rownum+1,57,row[3],eformat)#borrowers
    #blank out school district Libraries
    worksheet.write_blank('B6', None, eformat)#Beacon
    worksheet.write_blank('C6', None, eformat)#Beacon
    worksheet.write_blank('D6', None, eformat)#Beacon
    worksheet.write_blank('E6', None, eformat)#Beacon
    worksheet.write_blank('F6', None, eformat)#Beacon
    worksheet.write_blank('G6', None, eformat)#Beacon
    worksheet.write_blank('H6', None, eformat)#Beacon
    worksheet.write_blank('I6', None, eformat)#Beacon
    worksheet.write_blank('J6', None, eformat)#Beacon
    worksheet.write_blank('K6', None, eformat)#Beacon
    worksheet.write_blank('L6', None, eformat)#Beacon
    worksheet.write_blank('M6', None, eformat)#Beacon
    worksheet.write_blank('N6', None, eformat)#Beacon
    worksheet.write_blank('O6', None, eformat)#Beacon
    worksheet.write_blank('P6', None, eformat)#Beacon
    worksheet.write_blank('Q6', None, eformat)#Beacon
    worksheet.write_blank('R6', None, eformat)#Beacon
    worksheet.write_blank('S6', None, eformat)#Beacon
    worksheet.write_blank('T6', None, eformat)#Beacon
    worksheet.write_blank('U6', None, eformat)#Beacon
    worksheet.write_blank('V6', None, eformat)#Beacon
    worksheet.write_blank('W6', None, eformat)#Beacon
    worksheet.write_blank('X6', None, eformat)#Beacon
    worksheet.write_blank('Y6', None, eformat)#Beacon
    worksheet.write_blank('Z6', None, eformat)#Beacon
    worksheet.write_blank('AA6', None, eformat)#Beacon
    worksheet.write_blank('AB6', None, eformat)#Beacon
    worksheet.write_blank('AC6', None, eformat)#Beacon
    worksheet.write_blank('AD6', None, eformat)#Beacon
    worksheet.write_blank('AE6', None, eformat)#Beacon
    worksheet.write_blank('AF6', None, eformat)#Beacon
    worksheet.write_blank('AG6', None, eformat)#Beacon
    worksheet.write_blank('AH6', None, eformat)#Beacon
    worksheet.write_blank('AI6', None, eformat)#Beacon
    worksheet.write_blank('AJ6', None, eformat)#Beacon
    worksheet.write_blank('AK6', None, eformat)#Beacon
    worksheet.write_blank('AL6', None, eformat)#Beacon
    worksheet.write_blank('AM6', None, eformat)#Beacon
    worksheet.write_blank('AN6', None, eformat)#Beacon
    worksheet.write_blank('AO6', None, eformat)#Beacon
    worksheet.write_blank('AP6', None, eformat)#Beacon
    worksheet.write_blank('AQ6', None, eformat)#Beacon
    worksheet.write_blank('AR6', None, eformat)#Beacon
    worksheet.write_blank('AS6', None, eformat)#Beacon
    worksheet.write_blank('AT6', None, eformat)#Beacon
    worksheet.write_blank('AU6', None, eformat)#Beacon
    worksheet.write_blank('AV6', None, eformat)#Beacon
    worksheet.write_blank('AW6', None, eformat)#Beacon
    worksheet.write_blank('AX6', None, eformat)#Beacon
    worksheet.write_blank('AY6', None, eformat)#Beacon
    worksheet.write_blank('AZ6', None, eformat)#Beacon
    worksheet.write_blank('BA6', None, eformat)#Beacon
    worksheet.write_blank('BB6', None, eformat)#Beacon
    worksheet.write_blank('BC6', None, eformat)#Beacon
    worksheet.write_blank('BD6', None, eformat)#Beacon
    worksheet.write_blank('BE6', None, eformat)#Beacon
    worksheet.write_blank('BF6', None, eformat)#Beacon


    worksheet.write_blank('B16', None, eformat)#Clintondale
    worksheet.write_blank('C16', None, eformat)#Clintondale
    worksheet.write_blank('D16', None, eformat)#Clintondale
    worksheet.write_blank('E16', None, eformat)#Clintondale
    worksheet.write_blank('F16', None, eformat)#Clintondale
    worksheet.write_blank('G16', None, eformat)#Clintondale
    worksheet.write_blank('H16', None, eformat)#Clintondale
    worksheet.write_blank('I16', None, eformat)#Clintondale
    worksheet.write_blank('J16', None, eformat)#Clintondale
    worksheet.write_blank('K16', None, eformat)#Clintondale
    worksheet.write_blank('L16', None, eformat)#Clintondale
    worksheet.write_blank('M16', None, eformat)#Clintondale
    worksheet.write_blank('N16', None, eformat)#Clintondale
    worksheet.write_blank('O16', None, eformat)#Clintondale
    worksheet.write_blank('P16', None, eformat)#Clintondale
    worksheet.write_blank('Q16', None, eformat)#Clintondale
    worksheet.write_blank('R16', None, eformat)#Clintondale
    worksheet.write_blank('S16', None, eformat)#Clintondale
    worksheet.write_blank('T16', None, eformat)#Clintondale
    worksheet.write_blank('U16', None, eformat)#Clintondale
    worksheet.write_blank('V16', None, eformat)#Clintondale
    worksheet.write_blank('W16', None, eformat)#Clintondale
    worksheet.write_blank('X16', None, eformat)#Clintondale
    worksheet.write_blank('Y16', None, eformat)#Clintondale
    worksheet.write_blank('Z16', None, eformat)#Clintondale
    worksheet.write_blank('AA16', None, eformat)#Clintondale
    worksheet.write_blank('AB16', None, eformat)#Clintondale
    worksheet.write_blank('AC16', None, eformat)#Clintondale
    worksheet.write_blank('AD16', None, eformat)#Clintondale
    worksheet.write_blank('AE16', None, eformat)#Clintondale
    worksheet.write_blank('AF16', None, eformat)#Clintondale
    worksheet.write_blank('AG16', None, eformat)#Clintondale
    worksheet.write_blank('AH16', None, eformat)#Clintondale
    worksheet.write_blank('AI16', None, eformat)#Clintondale
    worksheet.write_blank('AJ16', None, eformat)#Clintondale
    worksheet.write_blank('AK16', None, eformat)#Clintondale
    worksheet.write_blank('AL16', None, eformat)#Clintondale
    worksheet.write_blank('AM16', None, eformat)#Clintondale
    worksheet.write_blank('AN16', None, eformat)#Clintondale
    worksheet.write_blank('AO16', None, eformat)#Clintondale
    worksheet.write_blank('AP16', None, eformat)#Clintondale
    worksheet.write_blank('AQ16', None, eformat)#Clintondale
    worksheet.write_blank('AR16', None, eformat)#Clintondale
    worksheet.write_blank('AS16', None, eformat)#Clintondale
    worksheet.write_blank('AT16', None, eformat)#Clintondale
    worksheet.write_blank('AU16', None, eformat)#Clintondale
    worksheet.write_blank('AV16', None, eformat)#Clintondale
    worksheet.write_blank('AW16', None, eformat)#Clintondale
    worksheet.write_blank('AX16', None, eformat)#Clintondale
    worksheet.write_blank('AY16', None, eformat)#Clintondale
    worksheet.write_blank('AZ16', None, eformat)#Clintondale
    worksheet.write_blank('BA16', None, eformat)#Clintondale
    worksheet.write_blank('BB16', None, eformat)#Clintondale
    worksheet.write_blank('BC16', None, eformat)#Clintondale
    worksheet.write_blank('BD16', None, eformat)#Clintondale
    worksheet.write_blank('BE16', None, eformat)#Clintondale
    worksheet.write_blank('BF16', None, eformat)#Clintondale


    worksheet.write_blank('B27', None, eformat)#Highland
    worksheet.write_blank('C27', None, eformat)#Highland
    worksheet.write_blank('D27', None, eformat)#Highland
    worksheet.write_blank('E27', None, eformat)#Highland
    worksheet.write_blank('F27', None, eformat)#Highland
    worksheet.write_blank('G27', None, eformat)#Highland
    worksheet.write_blank('H27', None, eformat)#Highland
    worksheet.write_blank('I27', None, eformat)#Highland
    worksheet.write_blank('J27', None, eformat)#Highland
    worksheet.write_blank('K27', None, eformat)#Highland
    worksheet.write_blank('L27', None, eformat)#Highland
    worksheet.write_blank('M27', None, eformat)#Highland
    worksheet.write_blank('N27', None, eformat)#Highland
    worksheet.write_blank('O27', None, eformat)#Highland
    worksheet.write_blank('P27', None, eformat)#Highland
    worksheet.write_blank('Q27', None, eformat)#Highland
    worksheet.write_blank('R27', None, eformat)#Highland
    worksheet.write_blank('S27', None, eformat)#Highland
    worksheet.write_blank('T27', None, eformat)#Highland
    worksheet.write_blank('U27', None, eformat)#Highland
    worksheet.write_blank('V27', None, eformat)#Highland
    worksheet.write_blank('W27', None, eformat)#Highland
    worksheet.write_blank('X27', None, eformat)#Highland
    worksheet.write_blank('Y27', None, eformat)#Highland
    worksheet.write_blank('Z27', None, eformat)#Highland
    worksheet.write_blank('AA27', None, eformat)#Highland
    worksheet.write_blank('AB27', None, eformat)#Highland
    worksheet.write_blank('AC27', None, eformat)#Highland
    worksheet.write_blank('AD27', None, eformat)#Highland
    worksheet.write_blank('AE27', None, eformat)#Highland
    worksheet.write_blank('AF27', None, eformat)#Highland
    worksheet.write_blank('AG27', None, eformat)#Highland
    worksheet.write_blank('AH27', None, eformat)#Highland
    worksheet.write_blank('AI27', None, eformat)#Highland
    worksheet.write_blank('AJ27', None, eformat)#Highland
    worksheet.write_blank('AK27', None, eformat)#Highland
    worksheet.write_blank('AL27', None, eformat)#Highland
    worksheet.write_blank('AM27', None, eformat)#Highland
    worksheet.write_blank('AN27', None, eformat)#Highland
    worksheet.write_blank('AO27', None, eformat)#Highland
    worksheet.write_blank('AP27', None, eformat)#Highland
    worksheet.write_blank('AQ27', None, eformat)#Highland
    worksheet.write_blank('AR27', None, eformat)#Highland
    worksheet.write_blank('AS27', None, eformat)#Highland
    worksheet.write_blank('AT27', None, eformat)#Highland
    worksheet.write_blank('AU27', None, eformat)#Highland
    worksheet.write_blank('AV27', None, eformat)#Highland
    worksheet.write_blank('AW27', None, eformat)#Highland
    worksheet.write_blank('AX27', None, eformat)#Highland
    worksheet.write_blank('AY27', None, eformat)#Highland
    worksheet.write_blank('AZ27', None, eformat)#Highland
    worksheet.write_blank('BA27', None, eformat)#Highland
    worksheet.write_blank('BB27', None, eformat)#Highland
    worksheet.write_blank('BC27', None, eformat)#Highland
    worksheet.write_blank('BD27', None, eformat)#Highland
    worksheet.write_blank('BE27', None, eformat)#Highland
    worksheet.write_blank('BF27', None, eformat)#Highland


    worksheet.write_blank('B39', None, eformat)#Mahopac
    worksheet.write_blank('C39', None, eformat)#Mahopac
    worksheet.write_blank('D39', None, eformat)#Mahopac
    worksheet.write_blank('E39', None, eformat)#Mahopac
    worksheet.write_blank('F39', None, eformat)#Mahopac
    worksheet.write_blank('G39', None, eformat)#Mahopac
    worksheet.write_blank('H39', None, eformat)#Mahopac
    worksheet.write_blank('I39', None, eformat)#Mahopac
    worksheet.write_blank('J39', None, eformat)#Mahopac
    worksheet.write_blank('K39', None, eformat)#Mahopac
    worksheet.write_blank('L39', None, eformat)#Mahopac
    worksheet.write_blank('M39', None, eformat)#Mahopac
    worksheet.write_blank('N39', None, eformat)#Mahopac
    worksheet.write_blank('O39', None, eformat)#Mahopac
    worksheet.write_blank('P39', None, eformat)#Mahopac
    worksheet.write_blank('Q39', None, eformat)#Mahopac
    worksheet.write_blank('R39', None, eformat)#Mahopac
    worksheet.write_blank('S39', None, eformat)#Mahopac
    worksheet.write_blank('T39', None, eformat)#Mahopac
    worksheet.write_blank('U39', None, eformat)#Mahopac
    worksheet.write_blank('V39', None, eformat)#Mahopac
    worksheet.write_blank('W39', None, eformat)#Mahopac
    worksheet.write_blank('X39', None, eformat)#Mahopac
    worksheet.write_blank('Y39', None, eformat)#Mahopac
    worksheet.write_blank('Z39', None, eformat)#Mahopac
    worksheet.write_blank('AA39', None, eformat)#Mahopac
    worksheet.write_blank('AB39', None, eformat)#Mahopac
    worksheet.write_blank('AC39', None, eformat)#Mahopac
    worksheet.write_blank('AD39', None, eformat)#Mahopac
    worksheet.write_blank('AE39', None, eformat)#Mahopac
    worksheet.write_blank('AF39', None, eformat)#Mahopac
    worksheet.write_blank('AG39', None, eformat)#Mahopac
    worksheet.write_blank('AH39', None, eformat)#Mahopac
    worksheet.write_blank('AI39', None, eformat)#Mahopac
    worksheet.write_blank('AJ39', None, eformat)#Mahopac
    worksheet.write_blank('AK39', None, eformat)#Mahopac
    worksheet.write_blank('AL39', None, eformat)#Mahopac
    worksheet.write_blank('AM39', None, eformat)#Mahopac
    worksheet.write_blank('AN39', None, eformat)#Mahopac
    worksheet.write_blank('AO39', None, eformat)#Mahopac
    worksheet.write_blank('AP39', None, eformat)#Mahopac
    worksheet.write_blank('AQ39', None, eformat)#Mahopac
    worksheet.write_blank('AR39', None, eformat)#Mahopac
    worksheet.write_blank('AS39', None, eformat)#Mahopac
    worksheet.write_blank('AT39', None, eformat)#Mahopac
    worksheet.write_blank('AU39', None, eformat)#Mahopac
    worksheet.write_blank('AV39', None, eformat)#Mahopac
    worksheet.write_blank('AW39', None, eformat)#Mahopac
    worksheet.write_blank('AX39', None, eformat)#Mahopac
    worksheet.write_blank('AY39', None, eformat)#Mahopac
    worksheet.write_blank('AZ39', None, eformat)#Mahopac
    worksheet.write_blank('BA39', None, eformat)#Mahopac
    worksheet.write_blank('BB39', None, eformat)#Mahopac
    worksheet.write_blank('BC39', None, eformat)#Mahopac
    worksheet.write_blank('BD39', None, eformat)#Mahopac
    worksheet.write_blank('BE39', None, eformat)#Mahopac
    worksheet.write_blank('BF39', None, eformat)#Mahopac

    worksheet.write_blank('B40', None, eformat)#Marlboro
    worksheet.write_blank('C40', None, eformat)#Marlboro
    worksheet.write_blank('D40', None, eformat)#Marlboro
    worksheet.write_blank('E40', None, eformat)#Marlboro
    worksheet.write_blank('F40', None, eformat)#Marlboro
    worksheet.write_blank('G40', None, eformat)#Marlboro
    worksheet.write_blank('H40', None, eformat)#Marlboro
    worksheet.write_blank('I40', None, eformat)#Marlboro
    worksheet.write_blank('J40', None, eformat)#Marlboro
    worksheet.write_blank('K40', None, eformat)#Marlboro
    worksheet.write_blank('L40', None, eformat)#Marlboro
    worksheet.write_blank('M40', None, eformat)#Marlboro
    worksheet.write_blank('N40', None, eformat)#Marlboro
    worksheet.write_blank('O40', None, eformat)#Marlboro
    worksheet.write_blank('P40', None, eformat)#Marlboro
    worksheet.write_blank('Q40', None, eformat)#Marlboro
    worksheet.write_blank('R40', None, eformat)#Marlboro
    worksheet.write_blank('S40', None, eformat)#Marlboro
    worksheet.write_blank('T40', None, eformat)#Marlboro
    worksheet.write_blank('U40', None, eformat)#Marlboro
    worksheet.write_blank('V40', None, eformat)#Marlboro
    worksheet.write_blank('W40', None, eformat)#Marlboro
    worksheet.write_blank('X40', None, eformat)#Marlboro
    worksheet.write_blank('Y40', None, eformat)#Marlboro
    worksheet.write_blank('Z40', None, eformat)#Marlboro
    worksheet.write_blank('AA40', None, eformat)#Marlboro
    worksheet.write_blank('AB40', None, eformat)#Marlboro
    worksheet.write_blank('AC40', None, eformat)#Marlboro
    worksheet.write_blank('AD40', None, eformat)#Marlboro
    worksheet.write_blank('AE40', None, eformat)#Marlboro
    worksheet.write_blank('AF40', None, eformat)#Marlboro
    worksheet.write_blank('AG40', None, eformat)#Marlboro
    worksheet.write_blank('AH40', None, eformat)#Marlboro
    worksheet.write_blank('AI40', None, eformat)#Marlboro
    worksheet.write_blank('AJ40', None, eformat)#Marlboro
    worksheet.write_blank('AK40', None, eformat)#Marlboro
    worksheet.write_blank('AL40', None, eformat)#Marlboro
    worksheet.write_blank('AM40', None, eformat)#Marlboro
    worksheet.write_blank('AN40', None, eformat)#Marlboro
    worksheet.write_blank('AO40', None, eformat)#Marlboro
    worksheet.write_blank('AP40', None, eformat)#Marlboro
    worksheet.write_blank('AQ40', None, eformat)#Marlboro
    worksheet.write_blank('AR40', None, eformat)#Marlboro
    worksheet.write_blank('AS40', None, eformat)#Marlboro
    worksheet.write_blank('AT40', None, eformat)#Marlboro
    worksheet.write_blank('AU40', None, eformat)#Marlboro
    worksheet.write_blank('AV40', None, eformat)#Marlboro
    worksheet.write_blank('AW40', None, eformat)#Marlboro
    worksheet.write_blank('AX40', None, eformat)#Marlboro
    worksheet.write_blank('AY40', None, eformat)#Marlboro
    worksheet.write_blank('AZ40', None, eformat)#Marlboro
    worksheet.write_blank('BA40', None, eformat)#Marlboro
    worksheet.write_blank('BB40', None, eformat)#Marlboro
    worksheet.write_blank('BC40', None, eformat)#Marlboro
    worksheet.write_blank('BD40', None, eformat)#Marlboro
    worksheet.write_blank('BE40', None, eformat)#Marlboro
    worksheet.write_blank('BF40', None, eformat)#Marlboro

    #Insert formulas to toal up all columns
    worksheet.write('A73', 'TOTAL', eformatlabel)#TOTAL
    worksheet.write_formula('B73', '=SUM(B2:B72)', eformat)#TOTAL
    worksheet.write_formula('C73', '=SUM(C2:C72)', eformat)#TOTAL
    worksheet.write_formula('D73', '=SUM(D2:D72)', eformat)#TOTAL
    worksheet.write_formula('E73', '=SUM(E2:E72)', eformat)#TOTAL
    worksheet.write_formula('F73', '=SUM(F2:F72)', eformat)#TOTAL
    worksheet.write_formula('G73', '=SUM(G2:G72)', eformat)#TOTAL
    worksheet.write_formula('H73', '=SUM(H2:H72)', eformat)#TOTAL
    worksheet.write_formula('I73', '=SUM(I2:I72)', eformat)#TOTAL
    worksheet.write_formula('J73', '=SUM(J2:J72)', eformat)#TOTAL
    worksheet.write_formula('K73', '=SUM(K2:K72)', eformat)#TOTAL
    worksheet.write_formula('L73', '=SUM(L2:L72)', eformat)#TOTAL
    worksheet.write_formula('M73', '=SUM(M2:M72)', eformat)#TOTAL
    worksheet.write_formula('N73', '=SUM(N2:N72)', eformat)#TOTAL
    worksheet.write_formula('O73', '=SUM(O2:O72)', eformat)#TOTAL
    worksheet.write_formula('P73', '=SUM(P2:P72)', eformat)#TOTAL
    worksheet.write_formula('Q73', '=SUM(Q2:Q72)', eformat)#TOTAL
    worksheet.write_formula('R73', '=SUM(R2:R72)', eformat)#TOTAL
    worksheet.write_formula('S73', '=SUM(S2:S72)', eformat)#TOTAL
    worksheet.write_formula('T73', '=SUM(T2:T72)', eformat)#TOTAL
    worksheet.write_formula('U73', '=SUM(U2:U72)', eformat)#TOTAL
    worksheet.write_formula('V73', '=SUM(V2:V72)', eformat)#TOTAL
    worksheet.write_formula('W73', '=SUM(W2:W72)', eformat)#TOTAL
    worksheet.write_formula('X73', '=SUM(X2:X72)', eformat)#TOTAL
    worksheet.write_formula('Y73', '=SUM(Y2:Y72)', eformat)#TOTAL
    worksheet.write_formula('Z73', '=SUM(Z2:Z72)', eformat)#TOTAL
    worksheet.write_formula('AA73', '=SUM(AA2:AA72)', eformat)#TOTAL
    worksheet.write_formula('AB73', '=SUM(AB2:AB72)', eformat)#TOTAL
    worksheet.write_formula('AC73', '=SUM(AC2:AC72)', eformat)#TOTAL
    worksheet.write_formula('AD73', '=SUM(AD2:AD72)', eformat)#TOTAL
    worksheet.write_formula('AE73', '=SUM(AE2:AE72)', eformat)#TOTAL
    worksheet.write_formula('AF73', '=SUM(AF2:AF72)', eformat)#TOTAL
    worksheet.write_formula('AG73', '=SUM(AG2:AG72)', eformat)#TOTAL
    worksheet.write_formula('AH73', '=SUM(AH2:AH72)', eformat)#TOTAL
    worksheet.write_formula('AI73', '=SUM(AI2:AI72)', eformat)#TOTAL
    worksheet.write_formula('AJ73', '=SUM(AJ2:AJ72)', eformat)#TOTAL
    worksheet.write_formula('AK73', '=SUM(AK2:AK72)', eformat)#TOTAL
    worksheet.write_formula('AL73', '=SUM(AL2:AL72)', eformat)#TOTAL
    worksheet.write_formula('AM73', '=SUM(AM2:AM72)', eformat)#TOTAL
    worksheet.write_formula('AN73', '=SUM(AN2:AN72)', eformat)#TOTAL
    worksheet.write_formula('AO73', '=SUM(AO2:AO72)', eformat)#TOTAL
    worksheet.write_formula('AP73', '=SUM(AP2:AP72)', eformat)#TOTAL
    worksheet.write_formula('AQ73', '=SUM(AQ2:AQ72)', eformat)#TOTAL
    worksheet.write_formula('AR73', '=SUM(AR2:AR72)', eformat)#TOTAL
    worksheet.write_formula('AS73', '=SUM(AS2:AS72)', eformat)#TOTAL
    worksheet.write_formula('AT73', '=SUM(AT2:AT72)', eformat)#TOTAL
    worksheet.write_formula('AU73', '=SUM(AU2:AU72)', eformat)#TOTAL
    worksheet.write_formula('AV73', '=SUM(AV2:AV72)', eformat)#TOTAL
    worksheet.write_formula('AW73', '=SUM(AW2:AW72)', eformat)#TOTAL
    worksheet.write_formula('AX73', '=SUM(AX2:AX72)', eformat)#TOTAL
    worksheet.write_formula('AY73', '=SUM(AY2:AY72)', eformat)#TOTAL
    worksheet.write_formula('AZ73', '=SUM(AZ2:AZ72)', eformat)#TOTAL
    worksheet.write_formula('BA73', '=SUM(BA2:BA72)', eformat)#TOTAL
    worksheet.write_formula('BB73', '=SUM(BB2:BB72)', eformat)#TOTAL
    worksheet.write_formula('BC73', '=SUM(BC2:BC72)', eformat)#TOTAL
    worksheet.write_formula('BD73', '=SUM(BD2:BD72)', eformat)#TOTAL
    worksheet.write_formula('BE73', '=SUM(BE2:BE72)', eformat)#TOTAL
    worksheet.write_formula('BF73', '=SUM(BF2:BF72)', eformat)#TOTAL

workbook.close()

#Create an email with an attachement
msg = MIMEMultipart()
msg['From'] = emailfrom
if type(emailto) is list:
    msg['To'] = ', '.join(emailto)
else:
    msg['To'] = emailto
msg['Date'] = formatdate(localtime = True)
msg['Subject'] = emailsubject
msg.attach (MIMEText(emailmessage))
part = MIMEBase('application', "octet-stream")
part.set_payload(open(excelfile,"rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition','attachment; filename=%s' % excelfile)
msg.attach(part)

#Send the email
smtp = smtplib.SMTP(emailhost, emailport)
#for Google connection
smtp.ehlo()
smtp.starttls()
smtp.login(emailuser, emailpass)
#end for Google connection
smtp.sendmail(emailfrom, emailto, msg.as_string())
smtp.quit()
