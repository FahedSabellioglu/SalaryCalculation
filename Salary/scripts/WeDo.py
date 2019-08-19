from Employee import Employee

"""
#parametreleri parametre2019.txt dosyasindan okuyabilmek icin fonksiyon:
def parametre_oku(line_number): 

    file = open('parametre2019.txt', mode = 'r') #parametre2019.txt ismindeki dosyayi okumak icin acar
    lines = file.read().split('\n') #dosyayi okuyup satirlara(line) boluyor
    words = lines[line_number].split(':') #satirlarda : ile ayrilmis kelimeleri boluyor


    lines[line_number]
    value = float(words[1]) #1. column'u yani sayisal degeri aliyor, float'a ceviriyor
    # 0. index: parametre ismi 1. index: parametre degeri

    if(file.close() == True):
        print("true dondu kapandi") #file kapanmiyor mu? bak bak bak

    return value

"""
# ----------------------------------------------------------------------------------------------------------------------------------


sale_5 = 0

convert_type = input("Do you want to calculate your net or gross salary? (N: Net, G: Gross) ")
if convert_type == 'N' or convert_type == 'n':
    gross_to_net = 1
    net_to_gross = 0

elif convert_type == 'G' or convert_type == 'g':
    net_to_gross = 1
    gross_to_net = 0

boss_cost = int(input("Do you want to see cost for employer (0: No, 1: Yes) "))
if boss_cost == 1:
    sale_5 = int(input(
        "Do we consider a 5% discount in the account of the employer's share of insurance premium? (0: No, 1: Yes) "))

marital_status = int(input("What is your marital status? (0: Single, 1: Married) "))
if marital_status == 1:
    partner_status = int(input("Does your partner work? (0: No, 1: Yes) "))
    kids_count = int(input("How many children do you have (0-...) "))
else:
    partner_status = 0
    kids_count = 0

# yeni bir Employee tanimliyor. isci primlerini, diger kesintilerini ve maasini hesaplamak icin
employee = Employee(kids_count, marital_status, partner_status, boss_cost, sale_5, gross_to_net, net_to_gross)
# hesaplamanin sonuclarini bulmak ve metin dosyasina yazdirmak icin bu fonksiyonu cagiriyor
# calculate fonksiyonunda gross_to_net_brut fonksiyonu cagiriliyor ve buradan gonderilen degerlere gore sonuclari .txt dosyasina kaydediyor.
employee.saveToFile()






