import openpyxl

filename = input('Type in .xlsx file you want to parse: ')
book = openpyxl.open(filename=filename)
sheet = book.active

amount = []
new_amount1, new_amount2 = [], []

prices = []
new_prices = []

dates = []

suppliers = []

clients_list = []


def dynamic_count():
    count = 0
    for row in sheet.iter_rows():
        count += 1

    return count


count_rows = dynamic_count()


def supply():
    print('\nsuppliers\n')
    for i in range(8, count_rows - 7 + 1, 4):
        if sheet[f'C{i}'].value is not None:
            anything = sheet[f'C{i}'].value.split('\n')
            supp = anything[-2].split()[0].split('-')[0].split('_')[0]

            match supp:
                case "НС":
                    supp = supp.replace('НС', 'Ерофеев')
                case "Т":
                    supp = supp.replace('Т', 'Курскпродукт')
                case "Э":
                    supp = supp.replace('Э', 'Эталон')
                case "З":
                    supp = supp.replace('З', 'Зернопродукт')

            suppliers.append(supp)
            print(supp)


# date
def second_column():
    print('\ndates\n')
    array2fill = [
        f'A{i}' for i in range(8, count_rows - 7 + 1, 4)
        if sheet[f'C{i}'].value is not None
    ]

    for j in range(len(array2fill)):
        print(sheet[array2fill[j]].value)
        dates.append(sheet[array2fill[j]].value)
        # remove this loop and parse data in new Excel file


# price to parse
def price_column():
    print('\nprices\n')
    for i in range(8, count_rows - 7 + 1, 4):
        if sheet[f'C{i}'].value is not None:
            strings = sheet[f'B{i}'].value.split()
            for string in strings:
                if string.startswith('Цена:'):
                    print(string[5:-1])
                    prices.append(string[5:-1])

    new_prices = [int(x) for x in prices]

    return new_prices


# amount to parse
def amount_column():
    print('\namount\n')
    for i in range(8, count_rows - 7 + 1, 4):
        if sheet[f'C{i}'].value is not None:
            strings = sheet[f'B{i}'].value.split()
            for string in strings:
                if string.startswith('Кол-во:'):
                    print(f"{string[-1]}{strings[strings.index(string) + 1]}")
                    amount.append(f"{string[-1]}{strings[strings.index(string) + 1]}")

    # in this function make a list of float numbers

    for el in amount:
        el = el.replace(',', '.')
        new_amount1.append(el)

    new_amount2 = list(map(float, new_amount1))

    return new_amount2


def clients():
    print('\nclients\n')
    for i in range(8, count_rows - 7 + 1, 4):
        if sheet[f'C{i}'].value is not None:
            print(sheet[f'B{i}'].value.split(',')[0])
            clients_list.append(sheet[f'B{i}'].value.split(',')[0])
    return clients_list


def multiplication_column():
    print('\ntotal cost\n')
    am_c = amount_column()
    p_c = price_column()
    for i in range(len(p_c)):
        print(am_c[i] * p_c[i])


def main():
    second_column()
    supply()
    clients()
    amount_column()
    price_column()
    # multiplication_column()


if __name__ == '__main__':
    print(f'\nInfo about {filename[:-5]} presented below')
    main()
