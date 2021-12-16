import openpyxl
from tqdm import tqdm


i = 0
f = open('model1.txt', 'r', encoding='utf-8')
f2 = open('model_word.txt', 'w', encoding='utf-8')


for line in f:
    # print(in_file)
    temp_string1 = ''.join([i for i in line if not i.isdigit()])
    temp_string2 = temp_string1.replace('.', '').replace('-', '').replace(' ', '').replace('e', '')
    # print(temp_string2)
    f2.write(temp_string2)
    i += 1
    if i % 100 == 0:
        print(i)


f.close()
f2.close()


def file_extract(path, savepath, file_name):
    print("파일 복사를 시작합니다 : " + file_name)
    MAX_VALUE = 50000

    excel_document = openpyxl.load_workbook(path + file_name + '.xlsx')

    sheet = excel_document.get_sheet_by_name('sheet')

    value = ['' for j in range(0, MAX_VALUE + 200)]

    for i in tqdm(range(2, MAX_VALUE)):
        value[i] = str(sheet.cell(row=i, column=17).value)  # .replace("\n", " ")
        i += 1

    f = open(savepath + file_name + ".txt", 'w', encoding='UTF-8')

    for i in tqdm(range(0, MAX_VALUE)):
        # data = " ".join(value[i].replace('\n ', '').split()) + "\n"
        data = value[i].replace('\n ', '').replace('\n\n', '\n').replace('\n\n', '\n').strip() + "\n"

        if data.__contains__('None'):
            break

        f.write(data)

    f.close()

