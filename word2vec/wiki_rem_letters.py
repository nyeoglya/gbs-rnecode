import re


i = 0
f = open('wiki_data_1.txt', 'r', encoding='utf-8')
f2 = open('wiki_data_1_out.txt', 'w', encoding='utf-8')


for line in f:
    # print(in_file)
    temp_string1 = ''.join([i for i in line if not i.isdigit()])
    temp_string2 = re.sub('[^0-9가-힣]', ' ', temp_string1)
    rr = re.sub('\n+', '', temp_string2)

    if rr == '\n':
        continue

    # print(' '.join(rr.split()))
    f2.write(' '.join(rr.split()) + '\n')
    i += 1
    if i % 100000 == 0:
        print(i)


f.close()
f2.close()
