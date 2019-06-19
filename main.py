from __future__ import print_function
import xlsxwriter
from suds.client import Client
from lxml import etree
from datetime import datetime
import apikey

APIKEY = apikey.APIKEY
WSDLFILE = 'https://datastore.ceidg.gov.pl/CEIDG.DataStore/services/DataStoreProvider201901.svc?wsdl'

client = Client(WSDLFILE,timeout=3600)

#Firmy z miast
cities = client.factory.create('ns1:ArrayOfstring')
cities.string.append('Wrocław')

#Kod PKD
pkd = client.factory.create('ns1:ArrayOfstring')
pkd.string.append('7711Z')

#Status 1 czyli aktywne firmy (Zgodnie z dokumentacją CEIDG)
status = client.factory.create('ns1:ArrayOfint')
status.int.append(1)

datefrom = datetime.strptime("2017-06-17","%Y-%m-%d")

#Wysyłamy zapytanie z parametrami
print("Zaczynam requestować do API")
result = client.service.GetMigrationData201901(AuthToken=APIKEY, PKD=pkd, Status=status, DataFrom=datefrom, City=cities)

#Konwertujemy otrzymany resultat na xml
print("Konwertuje na xml")
xmlParse = etree.fromstring(result)

#Tworzymi plik xlsx "daneFirmy", gdzie zostanie zapisana zawartość odczytana z CEIDG
summary = xlsxwriter.Workbook('daneFirmCEIDG.xlsx')
worksheet = summary.add_worksheet()

row = 0
print("Zapsisuje do excela")
#Iteruj po wynikach z CEIDG
for company in xmlParse.iter("InformacjaOWpisie"):
	#Podajemy co ma się zawierać w nazwie firmy

    try:
        # if(str(company[1][4].text).lower().__contains__("")):
        datefrom1 = datetime.strptime(company[4][0].text, "%Y-%m-%d")
        if (datefrom1 > datefrom):
            #nazwisko
            surname = company[1][1].text
            worksheet.write(row, 1, surname)
            #imię
            name = company[1][0].text
            worksheet.write(row, 2, name)
            #nazwa firmy
            companyName = company[1][4].text
            worksheet.write(row, 3, companyName)
            #adres email
            email = company[2][0].text
            worksheet.write(row, 4, email)
            #Strona www
            website = company[2][1].text
            worksheet.write(row, 5, website)
            #telefon
            phone = company[2][2].text
            worksheet.write(row, 6, phone)
            #miejscowość
            city = company[3][0][3].text
            worksheet.write(row, 7, city)
            #Województwo
            region = company[3][0][9].text
            worksheet.write(row, 8, region)  
            #Data rozpoczecia działalności
            dateFrom = company[4][0].text
            worksheet.write(row, 9, dateFrom)

            row += 1

    except:
        print("Nie udało się zapisać, wystąpił błąd!")



summary.close()

print("Program zakończył działanie")
