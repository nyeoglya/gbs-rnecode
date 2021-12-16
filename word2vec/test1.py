import gensim

# 사전 훈련된 Word2Vec 모델을 로드합니다.
model = gensim.models.KeyedVectors.load_word2vec_format('model.bin', binary=True)

print(model.vectors.shape)

# print(model.similarity('연어', '곰') * 100)
print(model.most_similar("국회의원"))

