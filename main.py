import os
import re
from datetime import datetime
import cx_Oracle
from openpyxl import Workbook
import ExcelWriter
from new_shortages import new_shortages



def purge(dir, pattern):
    for f in os.listdir(dir):
        if re.search(pattern, f):
            os.remove(os.path.join(dir, f))


# Set-up
connection = cx_Oracle.connect(
    "overseas4read",  # UID
    "OverseasSSW",  # PW
    "OTC-VS-018/ORAMULT"  # Server, SID implied Port ->1521
)
cursor = connection.cursor()

purge(os.getcwd(), '(.*)xlsx')
wb = Workbook()

cursor.execute(
    '''
SELECT nc.INTERNNUMMER,
       nc.OPDRACHTGEVER,
       nc.DOSSIERNUMMER_ERP,
       nc.DOCUMENT_REFERENTIENUMMER_MRN,
       nc.AANGIFTEDATUM
FROM NCTS.NCAANGH0 nc
WHERE ACTIEFBEDRIJF = 'ODC'
  AND nc.STATUSMESSAGE = 'REL_TRA'
ORDER BY nc.AANGIFTEDATUM
    '''
)

ws = wb.active
ws.title = 'NCTS-OPEN'
ExcelWriter.write_and_format(ws, cursor)

cursor.execute('''
SELECT d.MESSAGESTATUS, d.DECLARATIONID, d.COMMERCIALREFERENCE, d.ARC_AADREFERENCECODE, d.ISSUEDATE
FROM PLDA.ESDECLARATION d

WHERE ACTIVECOMPANY = 'ODC'
  AND d.MESSAGETYPE = '1'
  AND d.MESSAGESTATUS = 'ARC_ALL'
ORDER BY ISSUEDATE DESC

''')
ws = wb.create_sheet('EMCS-DEPARTURE OPEN')
ExcelWriter.write_and_format(ws, cursor)

cursor.execute('''
SELECT d.MESSAGESTATUS, d.DECLARATIONID, d.COMPLEMENTARYINFO, d.COMMERCIALREFERENCE,
 d.ARC_AADREFERENCECODE, d.ISSUEDATE, cl.SEQUENCE, cl.COMMERCIALDESC, cl.QUANTITY, cl.SHORTAGEOREXCESS, cl.OBSERVEDSHORTAGEEXCESS
FROM PLDA.ESDECLARATION d

LEFT JOIN PLDA.ESDECLARTARIF cl
ON d.DECLARATIONID = cl.DECLARATIONID
WHERE ACTIVECOMPANY = 'ODC'
  AND d.MESSAGETYPE = '1'
  AND d.MESSAGESTATUS = 'WRT_NOT' AND d.GLOBALCONCLUSIONRECEIPT = '2'
  AND cl.SHORTAGEOREXCESS <> ' '

ORDER BY ISSUEDATE DESC
''')

ws = wb.create_sheet('EMCS-DEPARTURE TEKORTEN')
ExcelWriter.write_and_format(ws, cursor)
new_shortages(ws)

cursor.execute('''
SELECT  cp.DECLARATION_ID, cp.PRINCIPAL, cp.COMMERCIALREFERENCE, cp.CUSTOMSMAINREFERENCENUMBER, cp.ISSUEDATE
        FROM PLDA.CPDECLARATION cp
        WHERE cp.ACTIVECOMPANY = 'ODC' AND cp.REGIMETYPE = 'A' AND cp.MESSAGESUBSTATUS <> 'WRT_NOT'
'''
               )
ws = wb.create_sheet('PLDA-EXPORT OPEN')
ExcelWriter.write_and_format(ws, cursor)

# Finish Up
str_time = datetime.now().strftime("%d_%m_%y")
wb.save(f'StreamReport{str_time}.xlsx')
connection.close()
