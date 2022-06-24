from datetime import datetime as dt
from openpyexcel import load_workbook
import csv
import math
import numpy as np


def today_is():
    return dt.now().strftime("%d.%m.%Y")

client = input("Кто Клиент ?")
cable = input("Примерное количество кабеля  в метрах?")
print("Вариант установки: ")

def variants():
    variant =  input("Аналог 2МП - 1 \nАналог 5МП - 2 \nАйпиш 2 МП - 3  \nАйпиш 5МП  - 4(в разработке) \n"" ")
    var_list = ["1","2","3","4"]
    if variant not in var_list:
        print(f"Дико извиняюсь, но вариантов  всего {len(var_list)}:")
        variants()
    else:
        return variant
variantus = variants()



def cam_calc_1():
    print("Вариант 1")
    quantity = input("Сколько нужно камер? ")
    if int(quantity) <= 0:
        print("Пардон, сударь, но, похоже, вы гоните")
        cam_calc_1()
    else:
        book = load_workbook(filename="Equipment.xlsx")
        sheet = book["Equip"]
        with open(f"Видеонаблюдение для {client.title()} на {today_is()}.csv", "w",encoding="utf-8",newline="") as file:
            writer = csv.writer(file)
            writer.writerow(("№","Наименование","Цена","количество","Сумма"))

        counter = 1
        list = [2,15,16]
        for i in list:
             with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",newline="") as file:
                 writer = csv.writer(file)
                 writer.writerow((counter,sheet[f"a{i}"].value,sheet[f"B{i}"].value,quantity,int(sheet[f"B{i}"].value)*int(quantity)))
                 counter+=1
        #Единичные товары(жесткий, рег)
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{17}"].value, sheet[f"B{17}"].value, cable, int(sheet[f"B{17}"].value) * int(cable)))
            counter+=1
            writer.writerow((counter, sheet[f"a{11}"].value, sheet[f"B{11}"].value, 1, int(sheet[f"B{11}"].value) * 1))
            counter+=1
            writer.writerow((counter, sheet[f"a{4}"].value, sheet[f"B{4}"].value, 1, int(sheet[f"B{4}"].value) * 1))
            counter += 1
        #подсчет суммы

        sum1 = 0
        sum2 = 0
        list2 = [4,10]
        for i in list:
            sum1 = int(sheet[f"B{i}"].value)*int(quantity)+sum1

        for i in list2:
            sum2 = int(sheet[f"B{i}"].value) + sum2

        # Блок питания :условие для подсчета количества БП
        power_supply = None

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
            #Формула количества блоков питания, PSQ
            psq = math.ceil(int(quantity)/7)
            print(psq)
            with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                      newline="") as file:
                writer = csv.writer(file)
                writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                                 int(sheet[f"B{13}"].value) * psq))
                power_supply = int(sheet[f"B{13}"].value) * psq




        #Всего кабеля
        sum_cab = int(cable)*int(sheet[f"B{16}"].value)



        #Итого
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(("","Итого","","", sum1+sum2+sum_cab+power_supply))



############################################################################################################

def cam_calc_2():
    print("Вариант 2")
    quantity = input("Сколько нужно камер? ")
    if int(quantity) <= 0:
        print("Пардон, сударь, но, похоже, вы гоните")
        cam_calc_2()
    else:
        book = load_workbook(filename="Equipment.xlsx")
        sheet = book["Equip"]
        with open(f"Видеонаблюдение для {client.title()} на {today_is()}.csv", "w", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow(("№", "Наименование", "Цена", "количество", "Сумма"))

        counter = 1
        list = [3, 14, 15, 17]
        for i in list:
            with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                      newline="") as file:
                writer = csv.writer(file)
                writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                                 int(sheet[f"B{i}"].value) * int(quantity)))
                counter += 1

        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow(
                (counter, sheet[f"a{16}"].value, sheet[f"B{16}"].value, cable, int(sheet[f"B{16}"].value) * int(cable)))
            counter += 1
            writer.writerow((counter, sheet[f"a{10}"].value, sheet[f"B{10}"].value, 1, int(sheet[f"B{10}"].value) * 1))
            counter += 1
            writer.writerow((counter, sheet[f"a{4}"].value, sheet[f"B{5}"].value, 1, int(sheet[f"B{5}"].value) * 1))
            counter += 1


        #Подсчёт сумм
        sum1 = 0
        sum2 = 0
        list2 = [5, 10]
        for i in list:

            sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

        for i in list2:
            sum2 = int(sheet[f"B{i}"].value) + sum2

        # Блок питания условие подсчета количества БП
        power_supply = None

        if int(quantity) < 2:
            with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                      newline="") as file:
                writer = csv.writer(file)
                writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                                 int(sheet[f"B{13}"].value) * 1))
                power_supply = int(sheet[f"B{13}"].value)
        elif int(quantity) >= 2 and int(quantity) < 4:
            with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                      newline="") as file:
                writer = csv.writer(file)
                writer.writerow((counter, sheet[f"a{12}"].value, sheet[f"B{12}"].value, 1,
                                 int(sheet[f"B{12}"].value) * 1))
                power_supply = int(sheet[f"B{12}"].value)
        else:
            # Формула количества блоков питания, PSQ

            psq = round(int(quantity) / 7.5)
            with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                      newline="") as file:
                writer = csv.writer(file)
                writer.writerow((counter, sheet[f"a{12}"].value, sheet[f"B{12}"].value, psq,
                                 int(sheet[f"B{12}"].value) * psq))
                power_supply = int(sheet[f"B{12}"].value) * psq

        # Всего кабеля

        sum_cab = int(cable) * int(sheet[f"B{16}"].value)

        # Итого
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(("Итого", "", "", "", sum1 + sum2 + sum_cab + power_supply))

#######################################################################################################################
def cam_calc_3():
    print("Вариант 3")
    quantity = input("Сколько нужно камер? ")
    if int(quantity) <= 0:
        print("Пардон, сударь, но, похоже, вы гоните")
        cam_calc_3()
    else:
        book = load_workbook(filename="Equipment.xlsx")
        sheet = book["Equip"]
        with open(f"Видеонаблюдение для {client.title()} на {today_is()}.csv", "w", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow(("№", "Наименование", "Цена", "количество", "Сумма"))

        counter = 1
        list = [8, 14, 15, 17]
        for i in list:
            with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                      newline="") as file:
                writer = csv.writer(file)
                writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                                 int(sheet[f"B{i}"].value) * int(quantity)))
                counter += 1
        # Единичные товары(жесткий, рег)
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow(
                (counter, sheet[f"a{16}"].value, sheet[f"B{16}"].value, cable, int(sheet[f"B{16}"].value) * int(cable)))
            counter += 1
            writer.writerow((counter, sheet[f"a{10}"].value, sheet[f"B{10}"].value, 1, int(sheet[f"B{10}"].value) * 1))
            counter += 1
            writer.writerow((counter, sheet[f"a{4}"].value, sheet[f"B{4}"].value, 1, int(sheet[f"B{4}"].value) * 1))
            counter += 1
        # подсчет суммы

        sum1 = 0
        sum2 = 0
        list2 = [4, 10]
        for i in list:
            sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

        for i in list2:
            sum2 = int(sheet[f"B{i}"].value) + sum2

        # Блок питания :условие для подсчета количества БП
        power_supply = None

        if int(quantity) < 2:
            with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                      newline="") as file:
                writer = csv.writer(file)
                writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                                 int(sheet[f"B{13}"].value) * 1))
                power_supply = int(sheet[f"B{13}"].value)
        elif int(quantity) >= 2 and int(quantity) < 4:
            with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                      newline="") as file:
                writer = csv.writer(file)
                writer.writerow((counter, sheet[f"a{12}"].value, sheet[f"B{12}"].value, 1,
                                 int(sheet[f"B{12}"].value) * 1))
                power_supply = int(sheet[f"B{12}"].value)
        else:
            # Формула количества блоков питания, PSQ
            psq = round(int(quantity) / 7.5)
            with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                      newline="") as file:
                writer = csv.writer(file)
                writer.writerow((counter, sheet[f"a{12}"].value, sheet[f"B{12}"].value, psq,
                                 int(sheet[f"B{12}"].value) * psq))
                power_supply = int(sheet[f"B{12}"].value) * psq

        # Всего кабеля
        sum_cab = int(cable) * int(sheet[f"B{16}"].value)

        # Итого
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(("Итого", "", "", "", sum1 + sum2 + sum_cab + power_supply))






def main():
    if variantus == "1":
        cam_calc_1()
    elif variantus == "2":
        cam_calc_2()
    # elif variants() == "3":
    #     cam_calc_3()
    # else:
    #     variants()
    print(f"Файл в формате сsv сформирован!")



if __name__ == "__main__":
    main()

#Условия :
# кабель  и варианты  в int или isnum()

# поправить порядок товаров

