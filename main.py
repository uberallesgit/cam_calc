from datetime import datetime as dt
from openpyexcel import load_workbook
import csv
import math
import numpy as np

########################################################################################################################
def today_is():
    return dt.now().strftime("%d.%m.%Y")
########################################################################################################################
client = input("Кто Клиент ?")
print(client)
if client == "":
    client ="somebody_someone"
cable = input("Примерное количество кабеля  в метрах?")
print(cable)
if not cable.isnumeric() :
    cable = "0"
else:
    print("cable = ",cable)

print("Вариант установки: ")


book = load_workbook(filename="Equipment.xlsx")
sheet = book["Equip"]
with open(f"Видеонаблюдение для {client.title()} на {today_is()}.csv", "w",encoding="utf-8",newline="") as file:
    writer = csv.writer(file)
    writer.writerow(("№","Наименование","Цена","количество","Сумма"))

########################################################################################################################
def variants():
    variant = input("Аналог 2МП - 1 "
                     "\nАналог 5МП - 2 "
                     "\nАйпиш 2 МП - 3  "
                     "\nАйпиш 5МП  - 4 "
                     "\nКомпак IP-Кам с SD - 5 \n")
    var_list = ["1","2","3","4","5"]
    if variant not in var_list:
        print(f"Дико извиняюсь, но вариантов  всего {len(var_list)}:")
        variants()
    else:
        return variant
########################################################################################################################
variants = variants()
########################################################################################################################
def how_many_cams():
    global quantity
    quantity = input("Сколько нужно камер? ")
    if int(quantity) <= 0:
        print("Пардон, сударь, но, похоже, вы гоните")
        how_many_cams()
    else:
        return quantity
########################################################################################################################

def cam_calc_1():
    print("Вариант 1...")
    counter = 1
    list = [2,15,16]
    for i in list:
         with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",newline="") as file:
             writer = csv.writer(file)
             writer.writerow((counter,sheet[f"a{i}"].value,sheet[f"B{i}"].value,quantity,int(sheet[f"B{i}"].value)*int(quantity)))
             counter+=1

    #Единичные товары(жесткий, БП)
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",newline="") as file:
        writer = csv.writer(file)
        #Считаем кабель
        writer.writerow((counter, sheet[f"a{17}"].value, sheet[f"B{17}"].value, cable, int(sheet[f"B{17}"].value) * int(cable)))
        counter+=1

        #Жесткий диск
        writer.writerow((counter, sheet[f"a{11}"].value, sheet[f"B{11}"].value, 1, int(sheet[f"B{11}"].value) * 1))
        counter+=1

    # Условие  выбора регистратора:

    reg_channels = input("Число каналов регистратора?  ")
    if reg_channels in ["0", " ", ""] : #or not reg_channels.isnumeric():
        reg_channels = quantity
    elif reg_channels < quantity:
        print("Число камер превышает  число портов регистратора, хотите изменить данные? \nДА = 1\nНЕТ = 0")
        res = input()
        if res in ["ДА", "да", "Да", "дА", "1", "хочу", "ага", "Давай", "Хочу", "yes", "Yes", "YES", "Мочи", "мочи"]:
            reg_channels = input("Число каналов регистратора ЕЩЁ РАЗ! ")

    if int(reg_channels) <= 4:
        global reg_count
        reg_count = 4

        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1

    elif 4 < int(reg_channels) <= 8:

        reg_count = 5
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1

    elif 8 < int(reg_channels) <= 16:

        reg_count = 6
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    elif 16 < int(reg_channels) <= 32:

        reg_count = 22
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    list2 = [reg_count, 11]

        # подсчет суммы



    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value)*int(quantity)+sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # Блок питания :условие для подсчета количества БП
    power_supply = 0

    if int(quantity) < 2:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif int(quantity) >= 2 and int(quantity) < 4 :
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
            power_supply = int(sheet[f"B{13}"].value)
    else:
        #расчет количества блоков питания (power suply quantity), PSQ
        psq = math.ceil(int(quantity)/7)
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
            power_supply = int(sheet[f"B{13}"].value) * psq




    #Всего кабеля
    sum_cab = int(cable)*int(sheet[f"B{17}"].value)



    #Итого
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("","Итого","","", sum1+sum2+sum_cab+power_supply))
########################################################################################################################

def cam_calc_2():
    print("Вариант 2...")

    counter = 1
    list = [3, 15, 16, 18]
    for i in list:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            counter += 1
    # Единичные товары(жесткий, БП)
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        # Считаем кабель
        writer.writerow(
            (counter, sheet[f"a{17}"].value, sheet[f"B{17}"].value, cable, int(sheet[f"B{17}"].value) * int(cable)))
        counter += 1
        # Жесткий диск
        writer.writerow((counter, sheet[f"a{11}"].value, sheet[f"B{11}"].value, 1, int(sheet[f"B{11}"].value) * 1))
        counter += 1

    # Условие  выбора регистратора:
    reg_channels = input("Число каналов регистратора?  ")
    if reg_channels in ["0", " ", ""] or not reg_channels.isnumeric():
        reg_channels = quantity
    elif reg_channels < quantity:
        print("Число камер превышает  число портов регистратора, хотите изменить данные? \nДА -1\nНЕТ - 0")
        res = input()
        if res in ["ДА", "да", "Да", "дА", "1", "хочу", "ага", "Давай", "Хочу", "yes", "Yes", "YES", "Мочи", "мочи"]:
            reg_channels = input("Число каналов регистратора ЕЩЁ РАЗ! ")

    if int(reg_channels) <= 4:
        global reg_count
        reg_count = 4

        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1

    elif 4 < int(reg_channels) <= 8:

        reg_count = 5
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1

    elif 8 < int(reg_channels) <= 16:

        reg_count = 6
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    elif 16 < int(reg_channels) <= 32:

        reg_count = 22
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    list2 = [reg_count, 11]

    # подсчет суммы

    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # Блок питания :условие для подсчета количества БП
    power_supply = 0

    if int(quantity) < 2:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif int(quantity) >= 2 and int(quantity) < 4:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
            power_supply = int(sheet[f"B{13}"].value)
    else:
        # расчет количества блоков питания (power suply quantity), PSQ
        psq = math.ceil(int(quantity) / 7)
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
            power_supply = int(sheet[f"B{13}"].value) * psq

    # Всего кабеля
    sum_cab = int(cable) * int(sheet[f"B{17}"].value)

    # Итого
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", sum1 + sum2 + sum_cab + power_supply))
########################################################################################################################

def ip_cam_calc_3():
    print("Вариант 3...")
    counter = 1
    list = [9, 21, 16, 18]
    for i in list:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            counter += 1
    # Единичные товары(жесткий, БП)
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)

        # Считаем кабель
        writer.writerow(
            (counter, sheet[f"a{24}"].value, sheet[f"B{24}"].value, cable, int(sheet[f"B{24}"].value) * int(cable)))
        counter += 1
        # Жесткий диск
        writer.writerow((counter, sheet[f"a{11}"].value, sheet[f"B{11}"].value, 1, int(sheet[f"B{11}"].value) * 1))
        counter += 1
    # Условие  выбора регистратора:
    reg_channels = input("Число каналов регистратора?  ")
    if reg_channels in ["0", " ", ""] or not reg_channels.isnumeric():
        reg_channels = quantity
    elif reg_channels < quantity:
        print("Число камер превышает  число портов регистратора, хотите изменить данные? \nДА -1\nНЕТ - 0")
        res = input()
        if res in ["ДА", "да", "Да", "дА", "1", "хочу", "ага", "Давай", "Хочу", "yes", "Yes", "YES", "Мочи", "мочи"]:
            reg_channels = input("Число каналов регистратора ЕЩЁ РАЗ! ")

    if int(reg_channels) <= 4:
        global reg_count
        reg_count = 4

        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    elif 4 < int(reg_channels) <= 8:

        reg_count = 5
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1

    elif int(reg_channels) == 9:

        reg_count = 23
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1

    elif 8 < int(reg_channels) <= 16:

        reg_count = 6
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    elif 16 < int(reg_channels) <= 32:

        reg_count = 22
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    list2 = [reg_count, 11]

    # подсчет суммы

    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # Блок питания :условие для подсчета количества БП
    power_supply = 0

    if int(quantity) < 2:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif int(quantity) >= 2 and int(quantity) < 4:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
            power_supply = int(sheet[f"B{13}"].value)
    else:
        # расчет количества блоков питания (power suply quantity), PSQ
        psq = math.ceil(int(quantity) / 7)
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
            power_supply = int(sheet[f"B{13}"].value) * psq

    # Всего кабеля
    sum_cab = int(cable) * int(sheet[f"B{24}"].value)

    # Итого
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", sum1 + sum2 + sum_cab + power_supply))
        print("sum1",sum1)
        print("sum2",sum2)
        print("sum_cab",sum_cab)
        print("power_supply",power_supply)
########################################################################################################################

def ip_cam_calc_4():
    print("Вариант 4...")

    counter = 1
    list = [10, 21, 16, 18]
    for i in list:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            counter += 1

    # Единичные товары(жесткий, БП)
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)

        # Считаем кабель
        writer.writerow(
            (counter, sheet[f"a{24}"].value, sheet[f"B{24}"].value, cable, int(sheet[f"B{24}"].value) * int(cable)))
        counter += 1
        # Жесткий диск
        writer.writerow((counter, sheet[f"a{11}"].value, sheet[f"B{11}"].value, 1, int(sheet[f"B{11}"].value) * 1))
        counter += 1

    # Условие  выбора регистратора:
    reg_channels = input("Число каналов регистратора?  ")
    if reg_channels in ["0", " ", ""] or not reg_channels.isnumeric():
        reg_channels = quantity
    elif reg_channels < quantity:
        print("Число камер превышает  число портов регистратора, хотите изменить данные? \nДА -1\nНЕТ - 0")
        res = input()
        if res in ["ДА", "да", "Да", "дА", "1", "хочу", "ага", "Давай", "Хочу", "yes", "Yes", "YES", "Мочи", "мочи"]:
            reg_channels = input("Число каналов регистратора ЕЩЁ РАЗ! ")
    if int(reg_channels) <= 4:
        global reg_count
        reg_count = 4

        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    elif 4 < int(reg_channels) <= 8:

        reg_count = 5
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1

    elif int(reg_channels) == 9:

        reg_count = 23
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1

    elif 8 < int(reg_channels) <= 16:

        reg_count = 6
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    elif 16 < int(reg_channels) <= 32:

        reg_count = 22
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    list2 = [reg_count, 11]

    # подсчет суммы

    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # Блок питания :условие для подсчета количества БП
    power_supply = 0

    if int(quantity) < 2:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif int(quantity) >= 2 and int(quantity) < 4:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
            power_supply = int(sheet[f"B{13}"].value)
    else:
        # расчет количества блоков питания (power suply quantity), PSQ
        psq = math.ceil(int(quantity) / 7)
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
            power_supply = int(sheet[f"B{13}"].value) * psq

    # Всего кабеля
    sum_cab = int(cable) * int(sheet[f"B{24}"].value)

    # Итого
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", sum1 + sum2 + sum_cab + power_supply))
        print("sum1", sum1)
        print("sum2", sum2)
        print("sum_cab", sum_cab)
        print("power_supply", power_supply)
########################################################################################################################

def ip_sd_calc_5():
    print("Вариант 5 ...")
    counter = 1
    list = [7,12,16]
    cam_sum = 0
    for i in list:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            cam_sum = (int(sheet[f"B{i}"].value) * int(quantity)) + cam_sum
            counter += 1

    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", cam_sum))
########################################################################################################################







def main():
    if variants == "1":
        how_many_cams()
        cam_calc_1()
    elif variants == "2":
        how_many_cams()
        cam_calc_2()
    elif variants == "3":
        how_many_cams()
        ip_cam_calc_3()
    elif variants == "4":
        how_many_cams()
        ip_cam_calc_4()
    elif variants == "5":
        how_many_cams()
        ip_sd_calc_5()
    else:
        variants()
    print(f"Файл в формате сsv сформирован!")
########################################################################################################################



if __name__ == "__main__":
    main()



