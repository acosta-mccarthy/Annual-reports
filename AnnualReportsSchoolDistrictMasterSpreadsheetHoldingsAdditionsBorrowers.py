#!/usr/bin/env python3

"""Create and email School district annual report on July 1st

Author: Nina Acosta
"""

import psycopg2
import xlsxwriter
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date, timedelta
yesterday = date.today() - timedelta(1)
year = yesterday.strftime ("%Y") #Pulls the year for the subject line and the body of the email

#SQL Query Part I & II - Holdings and Additions
q1='''
SELECT
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
--AS "eBook 2.13",
--AS "Audio Downloadable Units 2.17",
--AS "Total Videorecording Downloadable",
COUNT(CASE when icode2 = 'l' OR icode2 = 't' then 1 end)AS "Total Other Electronic Materials 2.19",
COUNT(CASE when icode2 = 'h' OR icode2 = 'r' then 1 end)AS "Total Sound Recording 2.21",
COUNT(CASE when icode2 = 'i' OR icode2 = 'q' then 1 end)AS "Total Videorecording 2.22",
COUNT(CASE when icode2 = 'f' OR icode2 = 'j' OR icode2 = 'm' OR icode2 = 's' then 1 end)AS "All Other Materials 2.23",
COUNT(CASE when icode2 = 'h' OR icode2 = 'r' OR icode2 = 'i' OR icode2 = 'q' OR icode2 = 'f' OR icode2 = 'j' OR icode2 = 'm' OR icode2 = 's' then 1 end)AS "Total Other Materials",
--AS "Grand Total Holdings",

--PART II - Total additions (all holdings added during previous year) by location and statistical code
COUNT(CASE when icode2 = 'a' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Fiction Added",
COUNT(CASE when icode2 = 'b' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Non-Fiction Added",
COUNT(CASE when icode2 = 'c' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Fiction Added",
COUNT(CASE when icode2 = 'd' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Non-Fiction Added",
COUNT(CASE when (icode2 = 'a' OR icode2 = 'b' OR icode2 = 'c' OR icode2 = 'd') AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Cataloged Books added 2.27",
COUNT(CASE when icode2 = 'l' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Software Added",
COUNT(CASE when icode2 = 't' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Software Added",
--AS "eBooks Added 2.29"
--AS "Electronic Materials Added"
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
-- AS "Other Media Added",
--AS "Downloadable Audio Added"
COUNT(CASE when (icode2 = 'n' OR icode2 = 'z') AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "All Other Print Materials Added 2.28"
-- AS "All other materials added"
-- AS "Total Added"


FROM
sierra_view.item_view

WHERE
location_code LIKE 'bea%' OR
location_code LIKE 'cld%' OR
location_code LIKE 'hil%' OR
location_code LIKE 'mah%' OR
location_code LIKE 'mar%'
--Limits locations to school district libraries only

GROUP BY "LOCATION"
ORDER BY "LOCATION"
'''

#SQL Query Part III - Borrowers
q2='''
SELECT

--PART III - Total borrowers by location and residency
LEFT(home_library_code, 3) AS "LOCATION",
COUNT(CASE when home_library_code IS NOT NULL AND ptype_code != '3' then 1 end)AS "Resident Borrowers 3.2",
COUNT(CASE when home_library_code IS NOT NULL AND ptype_code = '3' then 1 end)AS "Non-Resident Borrowers 3.3",
COUNT(CASE when home_library_code IS NOT NULL then 1 end)AS  "Total Number of Borrowers"

FROM
sierra_view.patron_view

WHERE
home_library_code LIKE 'bea%' OR
home_library_code LIKE 'cld%' OR
home_library_code LIKE 'hil%' OR
home_library_code LIKE 'mah%' OR
home_library_code LIKE 'mar%'
--Limits locations to school district libraries only

GROUP BY "LOCATION"
ORDER BY "LOCATION"
'''


#Name of Excel File
excelfile = "C:/Users/staff/Desktop/SDsampleReportForMAIUG_"+ str(year)+".xlsx" #Adds the year to the end of the filename

# These are variables for the email that will be sent.
# This code uses placeholders, please add your own email server info
emailhost = 'email.server.midhudson.org'
emailuser = 'emailaddress@midhudson.org'
emailpass = '*******'
emailport = '587'
emailsubject = 'School District Annual Report ' + str(year)
emailmessage = '''***This is an automated email***


The ''' + str(year) + ''' school district annual report masterlist is attached.
This spreadsheet contains the Holdings, Additions, and Borrowers for school district libraries.'''
emailfrom= 'emailaddress@midhudson.org'
emailto = 'nacosta@midhudson.org'


#This code uses placeholder info to connect to Sierra SQL server, please replace with your own info
conn = psycopg2.connect("dbname='iii' user='*****' host='000.000.000.000' port='1032' password='*****' sslmode='require'")



#Open session and run both queries
cursor = conn.cursor()
cursor.execute(q1)
rows = cursor.fetchall()
cursor.execute(q2)
lines = cursor.fetchall()
conn.close()

#Create Excel file
import xlsxwriter
workbook = xlsxwriter.Workbook("C:/Users/staff/Desktop/SDsampleReportForMAIUG_"+ str(year)+".xlsx")
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


    #Formula for Total Added Column [54]~OK~
    worksheet.write_formula('BC2', '=SUM(AJ2,AN2,BA2,BB2 )', eformat)#Total added
    worksheet.write_formula('BC3', '=SUM(AJ3,AN3,BA3,BB3 )', eformat)#Total added
    worksheet.write_formula('BC4', '=SUM(AJ4,AN4,BA4,BB4 )', eformat)#Total added
    worksheet.write_formula('BC5', '=SUM(AJ5,AN5,BA5,BB5 )', eformat)#Total added
    worksheet.write_formula('BC6', '=SUM(AJ6,AN6,BA6,BB6 )', eformat)#Total added

for rownum, row in enumerate(lines): #pulls borrower data from separate SQL query
    #worksheet.write(rownum+1,55,row[0],eformat)#Libraries
    worksheet.write(rownum+1,55,row[1],eformat)#Resident
    worksheet.write(rownum+1,56,row[2],eformat)#Non-resident
    worksheet.write(rownum+1,57,row[3],eformat)#borrowers

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
