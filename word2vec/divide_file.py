f = open('wiki_data_2_out.txt', 'r', encoding='utf-8')
lines = f.readline()

cnt = 1
filename_num = 1

for line in lines:
    fileName = "wiki_data_2_out_" + str(filename_num) + ".txt"

    fw = open(fileName, 'a', encoding='utf-8')
    fw.write(line)
    fw.close()

    if cnt == 10000000:
        print("파일 생성됨")
        fileName = filename_num + 1
        cnt = 0

    cnt = cnt + 1


f.close()
