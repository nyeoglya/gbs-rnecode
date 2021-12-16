import re


i = 0
f = open('wiki_data_1_out.txt', 'r', encoding='utf-8')
f2 = open('wiki_data_2_out.txt', 'w', encoding='utf-8')


for line in f:
    # print(in_file)
    rr = re.sub('\n+', '', line)

    if rr == '' or len(rr) < 10:
        continue

    f2.write(' '.join(rr.split()) + '\n')
    i += 1
    if i % 100000 == 0:
        print(i)


f.close()
f2.close()
