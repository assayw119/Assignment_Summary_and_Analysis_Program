from khaiii import KhaiiiApi
import docx2txt
import numpy as np
import pandas as pd
from collections import Counter
import re
import networkx as nx
import matplotlib
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib import rc

assingment = input('파일명을 입력해주세요 : ') # 2000000000_홍길동.docx
text = docx2txt.process(assingment)

api = KhaiiiApi()
sentence_list = []
morphs_list = []

df = pd.read_excel('stopwords.xlsx')
stopwords = list(df['불용어'])

document = text.split('\n')
for sentence in document:
    
    if sentence != '':
        # 모든 문장
        if '.' in sentence:
            sentence_detail = sentence.split('. ')
            for i in sentence_detail:
                if i != '':
                    sentence_list.append(i)
        else:
            sentence_list.append(sentence)


morphs_list = []
sentence_nouns_list = []

api = KhaiiiApi()

# 모든 문장에서 명사만 추출
for sentence in sentence_list:
    morphs_value = []
    for word in api.analyze(sentence):
        for morph in word.morphs:
            if morph.tag == 'NNG' and morph.lex not in stopwords: # 명사 추출, 불용어 제거
                morphs_list.append(morph.lex) # 문장별 단어 추출
                morphs_value.append(morph.lex) # 모든 단어 추출
    sentence_nouns_list.append(morphs_value) # 각 문장별 명사 리스트
    

num_top_nouns = 20 
morphs_counter = Counter(morphs_list)
morphs_top_nouns = dict(morphs_counter.most_common(num_top_nouns))


word2id = {w:i for i,w in enumerate(morphs_top_nouns.keys())}
id2word = {i:w for i,w in enumerate(morphs_top_nouns.keys())}

adjacent_matrix = np.zeros((num_top_nouns, num_top_nouns),int)
for sentence_nouns in sentence_nouns_list:
    for wi, i in word2id.items():
        if wi in sentence_nouns:
            for wj, j in word2id.items():
                if i != j and wj in sentence_nouns:
                    adjacent_matrix[i][j] += 1

adjacent_df = pd.DataFrame(adjacent_matrix, index=morphs_top_nouns.keys(), columns=morphs_top_nouns.keys())
adjacent_df.to_csv('test.csv')

word_network = nx.from_numpy_matrix(adjacent_matrix)


# font_path="./font/NanumBarunGothic.ttf"
# font_name = fm.FontProperties(fname=font_path).get_name()
# rc('font', family=font_name)
matplotlib.rcParams['axes.unicode_minus'] = False
rc('font', family='NanumBarunGothic')


fig = plt.figure()
plt.rcParams['font.family'] = 'AppleGothic'

plt.title(assingment)
fig.set_size_inches(20, 20)
ax = fig.add_subplot(1, 1, 1)
ax.axis("off")
option = {
    'node_color' : 'lightblue',
    'node_size' : 2000,
    'size' : 2
}
nx.draw(word_network, labels=id2word, ax=ax, font_family='AppleGothic', font_size=20)

plt.savefig('test.png')
plt.show()