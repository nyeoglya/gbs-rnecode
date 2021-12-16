from gensim.models.keyedvectors import KeyedVectors

model = KeyedVectors.load_word2vec_format('model1.bin', binary=True)
model.save_word2vec_format('model1.txt', binary=False)
