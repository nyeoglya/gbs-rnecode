import math
import time

import gensim

#  액셀 파일을 열기 위한 모듈
import openpyxl
from openpyxl import *

from tqdm import tqdm

from time import sleep


max_data = 2000

date_data = ['' for i in range(max_data)] # 데이터별 날짜
data = ['' for i in range(max_data)] # 모든 데이터를 저장할 변수
temp_data = ['' for i in range(max_data)]

excel = load_workbook("result_data.xlsx")
sheet = excel.active

cnt = 0
for row in sheet.rows:
    date_data[cnt] = str(row[0].value)
    data[cnt] = str(row[1].value)
    cnt += 1

print(cnt + " lines is detected")


# 사전 훈련된 Word2Vec 모델을 로드합니다.
model = gensim.models.KeyedVectors.load_word2vec_format('model1.bin', binary=True)

print(model.vectors.shape)

start = time.time()


# print(model.similarity('연어', '곰') * 100)
print(model.most_similar("국회의원"))

end = time.time()
print(f"{end - start:.5f} sec")
