#!/usr/bin/env python
#-*- coding: utf-8 -*-

import pyexcel as pe
from lxml import etree as ET

NS_XSI = "{http://www.w3.org/2001/XMLSchema-instance}"

#root = ET.XML('''\
#<?xml version="1.0"?>
#<!DOCTYPE root SYSTEM "test" [ <!ENTITY tasty "parsnips"> ]>
#<root>
#<a>&tasty;</a>
#</root>
#''')

root=ET.Element("eExact", )
root.set(NS_XSI + "noNamespaceSchemaLocation", "eExact-Schema.xsd")
#xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance",
#xsi:noNamespaceSchemaLocation="eExact-Schema.xsd")
xmlGLEntries = ET.SubElement(root, 'GLEntries')
xmlGLEntry = ET.SubElement(xmlGLEntries, 'GLEntry')
#xmlGLEntry.set("entry","59100014")
xmlGLEntry.set("status","E")
xmlDivision = ET.SubElement(xmlGLEntry, "Division")
xmlDivision.set("code", "001")
xmlDescription = ET.SubElement(xmlGLEntry, "Description")
xmlDescription.text = u'LP 01/06/ofp/2015'
xmlDate = ET.SubElement(xmlGLEntry, "Date")
xmlDate.text = u'2015-06-30'
xmlDocumentDate = ET.SubElement(xmlGLEntry, "DocumentDate")
xmlDocumentDate.text = '2015-06-30'
xmlJournal = ET.SubElement(xmlGLEntry, 'Journal')
xmlJournal.set("code","910")
xmlJournal.set("type","M")
xmlDescription2 = ET.SubElement(xmlJournal, "Description")
xmlDescription2.text = u'Listy płac'
xmlGLAccount = ET.SubElement(xmlJournal, "GLAccount", attrib={"code":"   999998","type":"B","subtype":"N","side":"D"})
xmlDescription3 = ET.SubElement(xmlGLAccount, "Description")
xmlDescription3.text = "Konto techniczne - bilansowe"
xmlAmmount = ET.SubElement(xmlGLEntry, 'Ammount')
xmlCurrency = ET.SubElement(xmlAmmount, "Currency", attrib={"code":"PLN"})
xmlValue = ET.SubElement(xmlAmmount, "Value")
xmlValue.text = "0"
xmlForeignAmount = ET.SubElement(xmlGLEntry, "ForeignAmount")
xmlCurrency2 = ET.SubElement(xmlForeignAmount, "Currency", attrib={"code":"PLN"})
xmlValue2 = ET.SubElement(xmlForeignAmount, "Value")
xmlValue2.text = "0"

xmlFinEntryLine = ET.SubElement(xmlGLEntry, "FinEntryLine", attrib={'subtype': 'C', 'type': 'N', 'number': '1'})
xmlDate = ET.SubElement(xmlFinEntryLine, 'Date' )
xmlDate.text = '2015-06-30'
xmlFinYear = ET.SubElement(xmlFinEntryLine,'FinYear',  attrib={'number': '2015'})
xmlFinPeriod  = ET.SubElement(xmlFinEntryLine,'FinPeriod', attrib={'number': '6'})
xmlGLAccount  = ET.SubElement(xmlFinPeriod,'GLAccount', attrib={'subtype': 'K', 'code': '   414001', 'type': 'W', 'side': 'D'})


xmlDescription4  = ET.SubElement(xmlGLAccount, "Description")
xmlDescription4.text = u'Wynagrodzenia z tytułu umowy o pracę'
xmlDescription5  = ET.SubElement(xmlFinPeriod, "Description")
xmlDescription5.text = 'LP 01/06/ofp/2015'
xmlCostcenter = ET.SubElement(xmlFinPeriod,"Costcenter", attrib={'code': 'SIEDLCE'})

xmlDescription6  = ET.SubElement(xmlCostcenter,"Description")
xmlDescription6.text = 'Siedlce'

'''
xmlGLAccount at 0x1023c3b48> GLAccount attrib={'subtype': 'J', 'code': '   999999', 'type': 'W', 'side': 'D'}

xmlDescription at 0x1023c3b90> Description {}
.text = 'Konto automatyczne - wynikowe'
xmlGLOffset at 0x1023c3b00> GLOffset attrib={'subtype': 'J', 'code': '   999999', 'type': 'W', 'side': 'D'}

xmlDescription at 0x1023c3b48> Description {}
.text = 'Konto automatyczne - wynikowe'
xmlCostunit at 0x1023c3c68> Costunit attrib={'code': 'I /GOSP.'}

xmlDescription at 0x1023c3b00> Description {}
.text = 'I/ Działal.gospodarcza'
xmlResource at 0x1023c3b90> Resource attrib={'number': '0'}
xmlItem at 0x1023c3c68> Item attrib={'code': 'H040'}

xmlDescription at 0x1023c3b48> Description {}
.text = 'Kierownik BT w Siedlcach'
xmlAssortment at 0x1023c3b90> Assortment attrib={'code': '0000', 'number': '0'}

xmlDescription at 0x1023c3c68> Description {}
.text = 'Miscellaneous'
xmlGLRevenue at 0x1023c3b00> GLRevenue attrib={'subtype': 'J', 'code': '   999999', 'type': 'W', 'side': 'D'}

xmlDescription at 0x1023c3b90> Description {}
.text = 'Konto automatyczne - wynikowe'
xmlGLCosts at 0x1023c3b48> GLCosts attrib={'subtype': 'J', 'code': '   999999', 'type': 'W', 'side': 'D'}

xmlDescription at 0x1023c3b00> Description {}
.text = 'Konto automatyczne - wynikowe'
xmlGLPurchase at 0x1023c3c68> GLPurchase attrib={'subtype': 'J', 'code': '   999999', 'type': 'W', 'side': 'D'}

xmlDescription at 0x1023c3b48> Description {}
.text = 'Konto automatyczne - wynikowe'
xmlWarehouse at 0x1023c3b90> Warehouse attrib={'code': '1'}

xmlDescription at 0x1023c3c68> Description {}
.text = 'Default warehouse'
xmlProject at 0x1023c3b00> Project attrib={'status': 'A', 'code': u'MI\u015a 2', 'type': 'I'}

xmlDescription at 0x1023c3b90> Description {}
.text = 'Projekt MIŚ aneks przedłużenie realizacji'
xmlSecurityLevel at 0x1023c3b48> SecurityLevel {}
.text = '10'
xmlDateStart at 0x1023c3b00> DateStart {}
.text = '2012-10-01'
xmlDateEnd at 0x1023c3b90> DateEnd {}
.text = '2014-12-31'
xmlAssortment at 0x1023c3c68> Assortment attrib={'number': '-1'}
xmlQuantity at 0x1023c3b48> Quantity {}
.text = '0'
xmlAmount at 0x1023c3b00> Amount {}

xmlCurrency at 0x1023c3b90> Currency attrib={'code': 'PLN'}
xmlDebit at 0x1023c3b48> Debit {}
.text = '4080'
xmlCredit at 0x1023c3c68> Credit {}
.text = '0'
xmlVAT at 0x1023c3b90> VAT attrib={'taxtype': 'V', 'vattype': 'N', 'code': '0', 'type': 'B'}

xmlDescription at 0x1023c3b48> Description {}
.text = 'Nie podlega ewidencji VAT'
xmlMultiDescriptions at 0x1023c3c68> MultiDescriptions {}

xmlMultiDescription at 0x1023c3b90> MultiDescription attrib={'number': '1'}
.text = 'VAT 0%'
xmlMultiDescription at 0x1023c3b48> MultiDescription attrib={'number': '2'}
.text = 'VAT 0%'
xmlMultiDescription at 0x1023c3c68> MultiDescription attrib={'number': '3'}
.text = 'VAT 0%'
xmlMultiDescription at 0x1023c3b90> MultiDescription attrib={'number': '4'}
.text = 'VAT 0%'
xmlPercentage at 0x1023c3b48> Percentage {}
.text = '0'
xmlCharged at 0x1023c3c68> Charged {}
.text = '0'
xmlVATExemption at 0x1023c3b00> VATExemption {}
.text = '0'
xmlExtraDutyPercentage at 0x1023c3b90> ExtraDutyPercentage {}
.text = '0'
xmlGLToPay at 0x1023c3c68> GLToPay attrib={'subtype': 'C', 'code': '   221200', 'type': 'B', 'side': 'C'}

xmlDescription at 0x1023c3b00> Description {}
.text = 'Podatek VAT należny'
xmlGLToClaim at 0x1023c3b48> GLToClaim attrib={'subtype': 'C', 'code': '   221100', 'type': 'B', 'side': 'D'}
.text = '

xmlDescription at 0x1023c3c68> Description {}
.text = 'Podatek VAT naliczony'
xmlCreditor at 0x1023c3b90> Creditor attrib={'code': '              000002', 'type': 'S', 'number': '000002'}

xmlName at 0x1023c3b48> Name {}
.text = 'Tax creditor'
xmlPaymentPeriod at 0x1023c3c68> PaymentPeriod {}
.text = 'M'
xmlVATListing at 0x1023c3b90> VATListing {}
.text = 'G'
xmlVATTransaction at 0x1023c3b48> VATTransaction attrib={'code': '0'}

xmlVATAmount at 0x1023c3b00> VATAmount {}
.text = '0'
xmlVATBaseAmount at 0x1023c3b90> VATBaseAmount {}
.text = '4080'
xmlVATBaseAmountFC at 0x1023c3c68> VATBaseAmountFC {}
.text = '4080'
xmlPayment at 0x1023c3b00> Payment {}

xmlPaymentMethod at 0x1023c3b90> PaymentMethod attrib={'code': 'B'}
xmlPaymentCondition at 0x1023c3b48> PaymentCondition attrib={'code': '00'}

xmlDescription at 0x1023c3b00> Description {}
.text = 'Gotówka'
xmlCSSDYesNo at 0x1023c3b90> CSSDYesNo {}
.text = 'K'
xmlCSSDAmount1 at 0x1023c3b48> CSSDAmount1 {}
.text = '0'
xmlCSSDAmount2 at 0x1023c3c68> CSSDAmount2 {}
.text = '0'
xmlInvoiceNumber at 0x1023c3b90> InvoiceNumber {}
.text = '10065135'
xmlDelivery at 0x1023c3b48> Delivery {}

xmlDate at 0x1023c3c68> Date {}
.text = '2015-07-08'
xmlFinReferences at 0x1023c3b00> FinReferences attrib={'TransactionOrigin': 'N'}

xmlUniquePostingNumber at 0x1023c3b48> UniquePostingNumber {}
.text = '0'
xmlYourRef at 0x1023c3c68> YourRef {}
.text = 'LP 01/06/ofp/2015'
xmlDocumentDate at 0x1023c3b00> DocumentDate {}
.text = '2015-06-30'
xmlDiscount at 0x1023c3b90> Discount {}

xmlPercentage at 0x1023c3c68> Percentage {}
.text = '0'
xmlPaymentTerms at 0x1023c3b48> PaymentTerms {}

xmlPaymentTerm at 0x1023c3b00> PaymentTerm attrib={'status': 'C', 'paymentMethod': 'T', 'paymentType': 'B', 'entry': '59100014', 'type': 'C', 'ID': '{5B7AF313-0FFA-42E5-A939-C012706A8F9F}', 'termType': 'W'}

xmlDescription at 0x1023c3c68> Description {}
.text = 'LP 01/06/ofp/2015'
xmlGLAccount at 0x1023c3b48> GLAccount attrib={'subtype': 'K', 'code': '   414001', 'type': 'W', 'side': 'D'}

xmlDescription at 0x1023c3b00> Description {}
.text = 'Wynagrodzenia z tytułu umowy o pracę'
xmlGLOffset at 0x1023c3b90> GLOffset attrib={'code': '   200100'}
xmlDebtor at 0x1023c3b48> Debtor attrib={'code': '                2347', 'number': '2347'}
xmlAmount at 0x1023c3c68> Amount {}

xmlCurrency at 0x1023c3b00> Currency attrib={'code': 'PLN'}
xmlDebit at 0x1023c3b48> Debit {}
.text = '0.0'
xmlCredit at 0x1023c3b90> Credit {}
.text = '118'
xmlVAT at 0x1023c3b00> VAT attrib={'code': '0'}
xmlForeignAmount at 0x1023c3b48> ForeignAmount {}

xmlCurrency at 0x1023c3b90> Currency attrib={'code': 'PLN'}
xmlDebit at 0x1023c3b00> Debit {}
.text = '0.0'
xmlCredit at 0x1023c3c68> Credit {}
.text = '118'
xmlRate at 0x1023c3b90> Rate {}
.text = '1'
xmlPaymentCondition at 0x1023c3b00> PaymentCondition attrib={'code': '14'}
xmlPercentage at 0x1023c3c68> Percentage {}
.text = '1'
xmlReference at 0x1023c3b90> Reference {}
.text = '2347/10065135'
xmlYourRef at 0x1023c3b48> YourRef {}
.text = 'LP 01/06/ofp/2015'
xmlInvoiceNumber at 0x1023c3b00> InvoiceNumber {}
.text = '10065135'
xmlInvoiceDate at 0x1023c3c68> InvoiceDate {}
.text = '2015-06-30'
xmlInvoiceDueDate at 0x1023c3b48> InvoiceDueDate {}
.text = '2015-07-14'
xmlProcessingDate at 0x1023c3b00> ProcessingDate {}
.text = '2015-07-12'
xmlReportingDate at 0x1023c3c68> ReportingDate {}
.text = '2015-06-30'
xmlResource at 0x1023c3b48> Resource attrib={'number': '0'}
xmlJournalization at 0x1023c3b00> Journalization {}

xmlResource at 0x1023c3b90> Resource attrib={'number': '0'}
xmlIsBlocked at 0x1023c3b48> IsBlocked {}
.text = '0'


'''

records = pe.iget_records(file_name="1.xlsx")
#book = pe.get_book(file_name="1.xlsx")

#sheet = book['Arkusz1']
#sheet.name_rows_by_column(0)
#print sheet.row[0]  #
#print sheet.column[u'DATA_SPRZEDAZY']
#print sheet


#print records

for record in records:
    #for i in record.keys():
        #print i
    print("%s * %s" % (record['KWOTA_BRUTTO'], record['STANOWISKO']))

tree = ET.parse("efka.xml")

#context = ET.iterparse("efka.xml")
#for action, elem in context:
#    print("%s: %s" % (action, elem.tag))
#    if not elem.text:
#            text = "None"
#        else:
#            text = elem.text
#        print elem.tag + " => " + text
#for child in tree.iter():
#    print("%s: %s" % (child.tag, child.text))
#    print tree.getpath(child), child.text, child.keys(), child.values(), child.tag
#    if len(child.keys()):
#        print child.attrib
#    a = child.getparent()
#    if a != None:
#        print(a.tail)
 #  'print child.xpath()

#print dir(xmlCurrency)

print ET.tostring(root, pretty_print=True, xml_declaration=True, encoding='UTF-8')

# write to file:
# tree = ET.ElementTree(root)
# tree.write('output.xml', pretty_print=True, xml_declaration=True)

#import elementtree.ElementTree as ET

# build a tree structure
#root = ET.Element("html")

#head = ET.SubElement(root, "head")

#title = ET.SubElement(head, "title")
#title.text = "Page Title"

#body = ET.SubElement(root, "body")
#body.set("bgcolor", "#ffffff")

#body.text = "Hello, World!"

# wrap it in an ElementTree instance, and save as XML
#tree = ET.ElementTree(root)#tree.write("page.xhtml")
