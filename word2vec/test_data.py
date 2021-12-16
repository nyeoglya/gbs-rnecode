import pandas as pd
import matplotlib.pyplot as plt
# import urllib.request
from gensim.models.word2vec import Word2Vec

from gensim.models import Word2Vec, KeyedVectors

import jpype

from konlpy.tag import Okt
from tqdm import tqdm

import sys

print(sys.version)

# urllib.request.urlretrieve("https://raw.githubusercontent.com/e9t/nsmc/master/ratings.txt", filename="ratings.txt")
train_data = pd.read_table('test_data.txt')
print("데이터의 상위 5개 값들을 불러옵니다")
print(train_data[:5])  # 상위 5개 출력

print("리뷰의 개수를 출력합니다" + str(len(train_data)))  # 리뷰 개수 출력

# NULL 값 존재 유무
print("널 값 존재 여부: " + str(train_data.isnull().values.any()))

train_data = train_data.dropna(how='any')  # Null 값이 존재하는 행 제거
# print(train_data.isnull().values.any())  # Null 값이 존재하는지 확인

print("널 값 제거한 데이터: " + str(len(train_data)))  # 리뷰 개수 출력

# 정규 표현식을 통한 한글 외 문자 제거
train_data['document'] = train_data['document'].str.replace("[^ㄱ-ㅎㅏ-ㅣ가-힣 ]", "")

print("한글 외 문자 제거, 5개 출력")
print(train_data[:5])  # 상위 5개 출력

# 불용어 정의
stopwords = ['의', '가', '이', '은', '들', '는', '좀', '잘', '걍', '과', '도', '를', '으로', '자', '에', '와', '한', '하다']

okt = Okt()

tokenized_data = []
for sentence in tqdm(train_data['document']):
    tokenized_sentence = okt.morphs(sentence, stem=True)  # 토큰화
    stopwords_removed_sentence = [word for word in tokenized_sentence if not word in stopwords]  # 불용어 제거
    tokenized_data.append(stopwords_removed_sentence)

print('데이터의 최대 길이 :', max(len(l) for l in tokenized_data))
print('데이터의 평균 길이 :', sum(map(len, tokenized_data)) / len(tokenized_data))
plt.hist([len(s) for s in tokenized_data], bins=50)
plt.xlabel('length of samples')
plt.ylabel('number of samples')
plt.show()


model = Word2Vec(sentences=tokenized_data, window=5, min_count=5, workers=4, sg=0)

# 완성된 임베딩 매트릭스의 크기 확인
print("임베딩 메트릭스 크기: " + str(model.wv.vectors.shape))

# print(model.wv.most_similar("전자"))

model.wv.save_word2vec_format('model1.bin', binary=True)
