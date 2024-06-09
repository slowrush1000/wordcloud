
# https://blog.naver.com/jazzlubu/223461247033
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import numpy as np
from PIL import *
# https://devpouch.tistory.com/196
import pandas as pd
#
import sys
import re
#
import seaborn as sns
from sklearn.feature_extraction.text import CountVectorizer
#
from collections import Counter
#
import networkx as nx

def my_token(text):
    tokens  = text.split()
    return tokens

class WordCloud4XLSX:
    def __init__(self):
        self.m_output_prefix    = ''
        self.m_filename         = ''
        self.m_column           = 0
        self.m_xlsx_datas       = []
        self.m_titles           = []
        self.m_lang             = 'kr'
        self.m_korean_font      = '/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf'
        self.m_persian_font     = '/usr/share/fonts/truetype/farsiweb/homa.ttf'
        self.m_korean_font_family   = 'NanumBarunGothic'
        self.m_persian_font_family  = 'homa'
    def PrintUsage(self):
        print(f'wordcloud4xlsx.py usage:')
        print(f'python3 wordcloud4xlsx.py output_prefix xlsx_file columns [kr|persian]')
    def ReadArgs(self, args):
        print(f'# read args start') 
        if 5 != len(args):
            self.PrintUsage()
            exit()
        self.m_output_prefix    = args[1]
        self.m_filename         = args[2]
        self.m_column           = args[3]
        if 'kr' == args[4]:
            self.m_lang         = 'kr'
        elif 'persian' == args[4]:
            self.m_lang         = 'persian'
        else:
            print(f'# error : {args[4]} is not defined keyword. You must set kr or persian!')
            exit()
        print(f'# read args end') 
    def PrintInputs(self):
        print(f'# print inputs start.')
        print(f'    output prefix : {self.m_output_prefix}')
        print(f'    xlsx file     : {self.m_filename}')
        print(f'    #column       : {self.m_column}')
        print(f'    lang          : {self.m_lang}')
        print(f'# print inputs end.')
    def ReadXlsx(self):
        print(f'# read xlsx file({self.m_filename}) start.')
        df = pd.read_excel(self.m_filename)
        titles  = None
        if 'kr' == self.m_lang:
            self.m_xlsx_datas  = df['제목']
        elif 'persian' == self.m_lang:
            self.m_xlsx_datas  = df['Title']
        #print(f'debug - {titles}')
        for title in self.m_xlsx_datas:
            temp_title  = str(title)
#            if 'kr' == self.m_lang:
#                temp_title  = self.ConvertStrOnlyKorean(temp_title)
#            elif 'persian' == self.m_lang:
#                temp_title  = self.ConvertStrOnlyPersian(temp_title)
            temp_title   = temp_title.strip()
            temp_title   = temp_title.replace(',', ' ')
            temp_title   = temp_title.replace('.', ' ')
            print(f'{temp_title}')
            self.m_titles.append(temp_title)
        print(f'# read xlsx file({self.m_filename}) end.')
    def MakeWordCloud(self):
        print(f'# make wordcloud start.')
        #
        total_titles    = ' '.join(self.m_titles)
        #print(f'debug - {total_titles}')
        #
        word_counter   = Counter(total_titles.split())
        #
        wordcloud   = None
        if 'kr' == self.m_lang:
            wordcloud = WordCloud(width=1000,
                                  height=1000,
                                  background_color='white',
                                  font_path=self.m_korean_font).generate_from_frequencies(word_counter)
        elif 'persian' == self.m_lang:
            wordcloud = WordCloud(width=1000,
                                  height=1000,
                                  background_color='white',
                                  font_path=self.m_persian_font).generate_from_frequencies(word_counter)
        plt.figure(figsize=(15,15))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.savefig(f'{self.m_output_prefix}.wordcloud.png')
        plt.show()
        print(f'# make wordcloud end.')
    def MakeNetwork(self):
        print(f'# make network start.')
        #
        cv          = CountVectorizer(ngram_range=(1,1), tokenizer=my_token)
        dtm_counts  = cv.fit_transform(self.m_titles)
        #dtm_df      = pd.DataFrame(dtm_counts.todense().tolist(), columns = cv.get_feature_names_out())
        #print(f'{dtm_df}')
        #
        dtm_counts  = (dtm_counts.T * dtm_counts)
        dtm_counts.setdiag(0)
        dtm_df      = pd.DataFrame(dtm_counts.todense().tolist(), columns = cv.get_feature_names_out(), index=cv.get_feature_names_out())
        print(f'{dtm_df}')
        # degree of centrality
        g           = nx.from_pandas_adjacency(dtm_df)
        dc          = nx.degree_centrality(g)
        dc_df       = sorted(dc.items(), key = lambda x: x[1], reverse = True)
        dc_df       = pd.DataFrame(list(dc_df)[:100], columns= ["word", "dc"])
        print(f'{dc_df}')
        # closeness of centrality
        cc          = nx.closeness_centrality(g)
        cc_df       = sorted(cc.items(), key = lambda x: x[1], reverse= True)
        cc_df       = pd.DataFrame(list(cc_df)[:100], columns = ["word", "cc"])
        print(f'{cc_df}')
        # betweeness of centrality
        bc          = nx.betweenness_centrality(g)
        bc_df = sorted(bc.items(), key = lambda x: x[1], reverse = True)
        bc_df = pd.DataFrame(list(bc_df)[:100], columns = ["word", "bc"])
        print(f'{bc_df}')
        # 
        temp            = pd.merge(dc_df, cc_df, how= "outer", on = "word")
        centrality_df   = pd.merge(temp, bc_df, how= "outer", on= "word")
        print(f'{centrality_df}')
        #
        print(f'debug #node : {g.number_of_nodes()}')
        print(f'debug #edge : {g.number_of_edges()}')
        edge_weight     = 9 # 9
        g_edge          = nx.Graph()
        g_edge.add_nodes_from(g.nodes(data = True))
        edges   = filter(lambda e:True if e[2]["weight"] >= edge_weight else False, g.edges(data = True))
        g_edge.add_edges_from(edges)
        for n in g.nodes():
            if len(list(nx.all_neighbors(g_edge, n))) == 0:
                g_edge.remove_node(n)
        print(f'debug #node : {g.number_of_nodes()}')
        print(f'debug #edge : {g.number_of_edges()}')
        plt.figure(figsize=(15,15))
        #
        centrality      = nx.degree_centrality(g_edge)
        #
        node_colors     = [centrality[node] for node in g_edge.nodes()]
        print(f'debug node colors min : {min(node_colors)}')
        #
        node_sizes      = [centrality[node] * 1500 for node in g_edge.nodes()]
        #
        pos             = nx.spring_layout(g_edge)
        nx.draw_networkx_nodes(g_edge, pos, node_color= node_colors, cmap= plt.cm.Reds, node_size= node_sizes, alpha= 0.7)
        nx.draw_networkx_edges(g_edge, pos, edge_color= 'grey', alpha= 0.5)
        #nx.draw_networkx_labels(g_edge, pos, font_family= 'NanumBarunGothic', font_size= 15)  # 라벨 폰트 크기 조정
        if 'kr' == self.m_lang:
            nx.draw_networkx_labels(g_edge, pos, font_family= self.m_korean_font_family, font_size= 10)  # 라벨 폰트 크기 조정
        elif 'persian' == self.m_lang:
            nx.draw_networkx_labels(g_edge, pos, font_family= self.m_persian_font_family, font_size= 10)  # 라벨 폰트 크기 조정
        # 중심성 값과 색상 맵의 대응 관계 표시
        #sm = plt.cm.ScalarMappable(cmap=plt.cm.Reds, norm=plt.Normalize(vmin=min(node_colors), vmax=max(node_colors)))
        #sm.set_array([])
        #plt.colorbar(sm, label="Degree Centrality")
        # 그래프의 바탕색 설정
        plt.gca().set_facecolor('#f0f0f0')  # 바탕색 설정
        plt.savefig(f'{self.m_output_prefix}.network.png')
        plt.show()
        print(f'# make network end.')
    def ConvertStr(self, text):
        text    = re.sub("[\x00-\x1F\x7F]", "", text)       # 제어문자 [[:cntrl:]]
        text    = re.sub("[\W]", " ", text)                 # 구두점 제거
        #text    = re.sub("[^\u3131-\u3163\uac00-\ud7a3]+", " ", text)  # 한글 이외의 문자 제거
        # text  = re.sub("(Zoom 수업|Zoom수업|줌 수업)", "줌수업", text)         # 2음절 단어
        # text  = re.sub("해지", "취소", text)               # 동음이의어
        return text
    def ConvertStrOnlyKorean(self, str):
        text    = re.sub("[\x00-\x1F\x7F]", "", text)       # 제어문자 [[:cntrl:]]
        text    = re.sub("[\W]", " ", text)                 # 구두점 제거
        text    = re.sub("[^\u3131-\u3163\uac00-\ud7a3]+", " ", text)  # 한글 이외의 문자 제거
        return text
    def ConvertStrOnlyPersian(self, str):
        return "".join(re.findall(r'[\u0600-\u06FF]+', str))
    def Run(self, args):
        self.ReadArgs(args)
        self.PrintInputs()
        self.ReadXlsx()
        self.MakeWordCloud()
        self.MakeNetwork()
    def TestWordCloud(self):
#        self.m_word_counter    = { "가가가"   : 10, "나나나"   : 5, "ccc"   : 2, "ddd"   : 1, }
        self.m_word_counter     = {'이란': 350, '대통령': 288, '헬기': 187, '사망': 117, '추락': 92, '라이시': 82, '확인': 42, '대통령,': 32, '추정': 30, '사망에': 28, '중동': 28, '속보': 27, '이란,': 27, '탑승': 26, '애도': 23, '비상착륙': 21, '탄': 21, '발견': 20, '정세': 20, '실종': 18, '사고': 17, '외무장관': 17, '최고지도자': 17, '잔해': 16, '구조대': 15, '공식': 15, '전원': 15, '수색': 14, '탑승자': 14, '9명': 13, '하메네이': 13, '생사': 12, '열원': 12, '등': 12, '권력': 11, '조전': 11, '급파': 9, '김정은,': 9, '추락으로': 8, '안': 8, '깊은': 8, '격랑': 8, '속': 8, '대선': 8, '차기': 8, '美': 8, '반지로': 8, '원인': 7, '추락사': 7, '2인자': 7, '정부': 7, '없어': 7, '중동정세': 7, '확산': 7, '이스라엘': 7, '사망,': 7, '악천후': 7, '헬기추락': 7, '기술적': 7, '신원': 7, '악천후에': 6, '난항': 6, '애도와': 6, '신호': 6, '잃은': 6, '혼란': 6, '불가피': 6, '제재': 6, '외교부': 6, '음모론': 6, '안갯속': 6, '후계자': 6, '지도자': 6, '강경파': 6, '왜': 6, '포착': 6, '총': 6, '핵': 6, '탓': 6, '美,': 6, '책임': 6, '국영통신': 6, '친근한': 6, '장례식': 6, '종합': 5, '헬기,': 5, '위로': 5, '대행': 5, '측근': 5, '정부,': 5, '당국자,': 5, '생존': 5, '영상': 5, '더': 5, '테헤란의': 5, '도살자': 5, '사고로': 5, '우려': 5, '튀르키예': 5, '드론,': 5, '돼': 5, '28일': 5, '충격적': 5, '미국,': 5, '것': 5, '접근': 4, '지연': 4, '불명': 4, '누구?': 4, '당국자': 4, '탄압': 4, '공백': 4, '대통령은': 4, '촉각': 4, '없다': 4, '요동': 4, '로이터': 4, '중': 4, '권력투쟁': 4, '추락,': 4, '전소': 4, '듯': 4, '6월': 4, 'vs': 4, '고장': 4, '심심한': 4, '미국': 4, '추모': 4, '美제재': 4, '손에': 4, '피': 4, '보궐선거': 4, '엄수': 4, '외부': 4, '총격': 4, '없었다': 4, '우라늄': 4, '안돼': 3, '아직': 3, '경착륙': 3, '언론': 3, '부통령,': 3, '모크베르': 3, '발표': 3, '차질': 3, '없이': 3, '운영': 3, '확인<로이터>': 3, '1968년': 3, '잇따라': 3, '가짜': 3, '내부': 3, '긴장': 3, '헬기는': 3, '속으로': 3, '수습': 3, '실시': 3, '꼬이나': 3, '대통령의': 3, '1순위': 3, '주도': 3, '승계': 3, '세계': 3, '거론': 3, '사고에': 3, '언급': 3, '마지막': 3, '모습': 3, '변화': 3, '후보': 3, '...': 3, '원인은': 3, '탑승자는': 3, '흔적': 3, '\\': 3, '아들': 3, '조선인민의': 3, '벗이었다': 3, '축하': 3, '내달': 3, '시작': 3, '테헤란': 3, '후계구도': 3, '불안': 3, '운집': 3, '하마스': 3, '참석': 3, '추락사고': 3, '수': 3, '핵개발': 3, '비난': 3, 'IAEA': 3, '에브라힘': 2, '현장에': 2, '작업': 2, '급파,': 2, '미확인': 2, '희망': 2, '동승': 2, '상보': 2, '국경': 2, '혁명수비대': 2, '적신월사': 2, '직무': 2, '국민의': 2, '종': 2, '국정': 2, '<로이터>': 2, '초도비행한': 2, '미국산': 2, '기종': 2, '국제사회': 2, '반체제': 2, '숙청': 2, '강경': 2, '보수파': 2, '안갯속으로': 2, '강경책': 2, '출신': 2, '라이시의': 2, '초강경': 2, '넘은': 2, '미국제': 2, '벨-212': 2, '무관': 2, '헬기사고': 2, 'SNS': 2, '정비': 2, '가능성': 2, '끝': 2, '시신': 2, '이내': 2, '장관': 2, '선포': 2, '강경보수': 2, '적자': 2, '행보': 2, '살얼음판': 2, '내우외환': 2, '땐': 2, '대혼란': 2, '슬픔': 2, '극복하길': 2, '전해': 2, '요동칠까': 2, '주목': 2, '유혈': 2, '차기지도자': 2, '솔솔': 2, '직대': 2, '구도': 2, '긴장감': 2, '고조되는': 2, '민심': 2, '파장은?': 2, '중단': 2, '공식확인': 2, '다시': 2, '모두': 2, '들썩이는': 2, '시장': 2, '사우디': 2, '국왕은': 2, '건강악화': 2, '직전': 2, '7월': 2, '세습': 2, '소식에': 2, '서방': 2, '곧': 2, '성명': 2, '부통령': 2, '핫이슈': 2, '외무장관,': 2, '누구': 2, '태운': 2, '전망': 2, '-로이터': 2, '모든': 2, '당국,': 2, '전': 2, '시': 2, '중국': 2, '지원': 2, '날씨': 2, '못해': 2, '넘게': 2, '못': 2, '잔해추정': 2, '국내외': 2, '동승자': 2, '안개': 2, '후': 2, '추락한': 2, '악천후로': 2, '총력전': 2, '이란대통령': 2, '찾아': 2, '추락에': 2, '오리무중': 2, '생사확인': 2, '소식,': 2, '외교': 2, '경제': 2, '미제': 2, '띄운': 2, '건': 2, '결함': 2, '미': 2, '김정은': 2, '죽음': 2, '춤추고': 2, '청년들': 2, '인권': 2, '은밀한': 2, '요동치는': 2, '중도': 2, '사망에도': 2, '국제유가': 2, '푸틴': 2, '묻혀': 2, '미,': 2, '묻힌': 2, '정치': 2, '추모객': 2, '조짐': 2, '6월28일': 2, '분열': 2, '정권': 2, '일축': 2, '고장으로': 2, '벗': 2, '충격': 2, '우방': 2, '아냐': 2, '테헤란서': 2, '장례식에': 2, '수백만명': 2, '지도자도': 2, '조기': 2, '권력승계': 2, '공격': 2, '조사결과': 2, '같은': 2, '알아볼': 2, '신원,': 2, '형체': 2, '결의': 2, '압박': 2, '20명': 2, '등록': 2, '합의': 2, '고농축': 2, '증가': 2, '확인되지': 1, '않아': 1, '(종합)': 1, '휘말려': 1, '피해상황': 1, '실종...': 1, '일행': 1, '경착륙...현장': 1, '급파(상보)': 1, '위기,': 1, '버려': 1, '아제르바이잔': 1, '인근': 1, '<이란': 1, '언론>': 1, '인원은': 1, '대통령은?': 1, '후계자이자': 1, '사망...이란': 1, '생명': 1, '<타스>': 1, '아주경제': 1, '오늘의': 1, '뉴스': 1, '外': 1, '직무대행': 1, '모크베르는': 1, '사후': 1, '기대': 1, '낮아': 1, '노후': 1, '가능성↑': 1, '...대통령': 1, '평안한': 1, '안식을': 1, '부통령이': 1, '권력공백': 1, '인사': 1, '히잡시위': 1, '2위': 1, '허위정보': 1, 'SNS에': 1, '거짓': 1, '200만뷰': 1, '넘어': 1, '과거': 1, '사진으로': 1, '루머도': 1, '결속하려': 1, '택할': 1, '수도': 1, '주도한': 1, '검사': 1, '랍비들': 1, '죽음은': 1, '신의': 1, '응징': 1, '초대형': 1, '변수에': 1, '때린': 1, '추락사...': 1, '숨죽인': 1, '55년': 1, '측': 1, '죽음과': 1, '우리는': 1, '영상?': 1, '확인?': 1, '가짜뉴스': 1, '범람': 1, '때문?': 1, '어려움': 1, '표한다': 1, '드론이': 1, '총력': 1, '불량': 1, '가능성도': 1, '맡은': 1, '바게리': 1, '차관은': 1, '임시': 1, '대통령승인...': 1, '50일': 1, '韓정부,': 1, '하메네이,': 1, '5일간': 1, '애도기간': 1, '정적': 1, '소행?': 1, '관여?...': 1, '난무': 1, '반서구': 1, '반인권': 1, '공백까지': 1, '다툼': 1, '좋은': 1, '친구': 1, '잃었다,': 1, '中': 1, '사고,': 1, '위로의': 1, '뜻': 1, '反이스라엘': 1, '선봉': 1, '가자': 1, '휴전': 1, '협상': 1, '尹정부': 1, '기원': 1, '여적': 1, '히잡': 1, '시위': 1, '진압한': 1, '공백에': 1, '커질듯': 1, '가자전쟁': 1, '장기화에': 1, '엎친데': 1, '덮친격': 1, '보수강경파': 1, '별명도': 1, '라': 1, '불린': 1, '보수': 1, '노선': 1, '휩싸인': 1, '무게': 1, '일각선': 1, '모크베르도': 1, '국가애도': 1, '기간': 1, '제1부통령,': 1, '다툼,': 1, '이반': 1, '배후설': 1, '관리': 1, '우리와': 1, '법조인': 1, '최고지도자로': 1, '시험대': 1, '오른': 1, '갑작스런': 1, '정국': 1, '대미': 1, '공백,': 1, '갑자기': 1, '작고한': 1, '과거,': 1, '표적': 1, '삼을': 1, '이유': 1, '1대만': 1, '타살': 1, '음모론도': 1, '증권거래소': 1, '거래': 1, '이란정부,': 1, '뉴스속': 1, '용어': 1, '핵합의': 1, '재추진': 1, '관리,': 1, '헬리콥터': 1, '관여': 1, '해': 1, '대통령-외무': 1, '나라': 1, '안팎에': 1, '전역서': 1, '전소된': 1, '채': 1, '에': 1, '금': 1, '구리값': 1, '역대': 1, '최고': 1, '유가': 1, '영향받을까': 1, '과': 1, '드론으로': 1, '본': 1, '현장': 1, '사진까지': 1, '덮친': 1, '출렁이나': 1, '예기치': 1, '못한': 1, '죽음에...각국': 1, '복잡한': 1, '거부감': 1, '아들을': 1, '최고지도자로?': 1, '3대': 1, '그': 1, '헬기만': 1, '두고': 1, '열리나': 1, '아들에': 1, '내각': 1, '사망...': 1, '앞으로': 1, '올까?': 1, '금값': 1, '장중': 1, '2450달러': 1, '술렁이는': 1, '위기의': 1, 'EU': 1, '경계': 1, '불보듯': 1, '파장에': 1, '요동치나': 1, '내각,': 1, '긴급': 1, '회의': 1, '소집...정부': 1, '이륙까지는': 1, '기상': 1, '좋았는데': 1, '숨진': 1, '손꼽던': 1, '현지매체': 1, '추락..전원': 1, '생존자': 1, '포함,': 1, '숨져': 1, '이란당국': 1, '고조될': 1, '듯(종합)': 1, '고개': 1, '드는': 1, '순간...': 1, '댐': 1, '준공식': 1, '다녀오다': 1, '만났나': 1, '2보': 1, '추정...부통령': 1, '내부서': 1, '사망설': 1, '흘러나와': 1, '전소,': 1, '고위관리': 1, '(상보)': 1, '사망한': 1, '추락사고로': 1, '가짜인데': 1, '160만뷰': 1, '허위': 1, '정보': 1, '포함': 1, '위한': 1, '희생양?': 1, '특유의': 1, '구조는': 1, '유력': 1, '공개돼': 1, '위치': 1, '<스푸트니크>': 1, '후계자였는데': 1, '공석': 1, '초도': 1, '비행': 1, '구출에': 1, '제공': 1, '헬기추락에': 1, '우려,': 1, '예의주시중': 1, '잔해?': 1, '위치에': 1, '실종된': 1, '불투명': 1, '공석은?': 1, '러시아': 1, '타스': 1, '지역서': 1, '탑승,': 1, '구조팀': 1, '12시간': 1, '찾는': 1, '이유는?': 1, '9명,': 1, '기체': 1, '...생사확인': 1, '생사여부': 1, '유고시,': 1, '정책': 1, '변화는': 1, '확인...좌표': 1, '공유': 1, '추위로': 1, '...탑승자는': 1, '슬라이드': 1, '포토': 1, '안개가': 1, '자욱...': 1, '사라진': 1, '전역': 1, '파장': 1, '러': 1, 'EU,': 1, '승무원': 1, '핸드폰서': 1, '17시간': 1, '추락...최고지도자': 1, '기도해야': 1, '사망시': 1, '승계는?': 1, '美월간': 1, '더애틀란틱': 1, '반경': 1, '2㎞': 1, '개발': 1, '무장세력': 1, '지원,시위': 1, '진압': 1, '지점': 1, '바이든도': 1, '보고': 1, '받아': 1, '탑승한': 1, '하마스,': 1, '완전한': 1, '연대': 1, '표명': 1, '추락...': 1, '난항,': 1, '확인,': 1, '정확한': 1, '위치는': 1, '찾지': 1, '외무장관과': 1, '동승한': 1, '실종...수색': 1, '행방': 1, '10시간': 1, '지나도': 1, '군,': 1, '현장으로': 1, '향해': 1, '유럽,': 1, '보도': 1, '예의주시': 1, '눈보라': 1, '난항으로': 1, '상황': 1, '면밀히': 1, '주시': 1, '외무장관도': 1, '실종...악천후가': 1, '원인인': 1, '동승\\': 1, '인근서': 1, '짙은': 1, '안개에': 1, '시야': 1, '제한': 1, '이동': 1, '비상착륙...이란': 1, '매체,': 1, '기도': 1, '방송': 1, '선거': 1, '신학': 1, '물망': 1, '위': 1, '최고지도자가': 1, '정점': 1, '통치구조는?': 1, '안전': 1, '문제는': 1, '그들': 1, '첫': 1, '추락이': 1, '핵위기': 1, '고조시키는': 1, '이유세모금': 1, '김정은도': 1, '최악인데': 1, '벼랑': 1, '내몰린': 1, '이란의': 1, '미래는?': 1, '이슈&뷰': 1, '北김정은,': 1, '50년': 1, '탓?': 1, '미국탓': 1, '45년된': 1, '해운': 1, '석유株': 1, '기술': 1, '제재가': 1, '불러': 1, '지지자들': 1, '전국적': 1, '물결': 1, '은': 1, '호전성': 1, '강화?': 1, '신중론': 1, '유지?': 1, '향후': 1, '노선에': 1, '벌써': 1, '놀란': 1, '암운': 1, '세습통치': 1, '부활하나': 1, '독재자의': 1, '사망했는데,': 1, '환호한': 1, '헬기추락은': 1, '기술결함': 1, '원인에': 1, '지목': 1, '한편에선': 1, '1인자': 1, '아들,': 1, '급부상': 1, '후보군에': 1, '개혁파': 1, '포함될까': 1, '후계': 1, '이어': 1, '사우디도': 1, '패권국': 1, '흔들': 1, '유엔': 1, '안보리서': 1, '묵념...이스라엘': 1, '거센': 1, '반발': 1, '자욱한': 1, '날,': 1, '낡은': 1, '헬기를…': 1, '선택': 1, '왜?': 1, '추락사에': 1, '노후헬기': 1, '띄운건': 1, '장례일정': 1, '23일': 1, '고향에': 1, '안장': 1, '말고': 1, '순교': 1, '라고': 1, '표현하라': 1, '나선': 1, '내림세': 1, '지속': 1, '주장도': 1, '막후': 1, '실세?': 1, '거론되는': 1, '모즈타바': 1, '하메네이는': 1, '망가진': 1, '신정국가': 1, '대선에': 1, '국운': 1, '달려': 1, '사망했는데': 1, '불꽃놀이': 1, '환호,': 1, '국무부': 1, '애도하나,': 1, '많은': 1, '엇갈린': 1, '`기술적': 1, '고장`으로': 1, '추락...미국': 1, '보내': 1, '애도하며': 1, '사람': 1, '충격,': 1, '대규모': 1, '추락死': 1, '장례식,': 1, '내일부터': 1, '도시': 1, '옮겨가며': 1, '진행': 1, '지적': 1, '에르도안': 1, '협력': 1, '강화': 1, '놓고': 1, '충돌땐': 1, '때문에': 1, '죽었다': 1, '美책임론': 1, '나온': 1, '이유는?핫이슈': 1, '장례': 1, '사흘간': 1, '죽음에': 1, '사회': 1, '심화...': 1, '바뀌지': 1, '않는': 1, '게': 1, '슬퍼': 1, '보선': 1, '6월8일': 1, '온건파': 1, '설자리': 1, '묻은': 1, '사실은': 1, '불변': 1, '기적적으로': 1, '탈출해': 1, '가짜영상': 1, '200만명': 1, '낚였다': 1, '미국은': 1, '아무': 1, '역할': 1, '했다': 1, '개입': 1, '의혹': 1, '블랙스완': 1, '저자': 1, '이슈,': 1, '투자자들은': 1, '우려할': 1, '필요': 1, '추락헬기,': 1, '여파로': 1, '부품을': 1, '암시장에서': 1, '조달': 1, '원인,': 1, '공방': 1, '국영통신,': 1, '원인으로': 1, '애도했지만': 1, '많이': 1, '자': 1, '착수...美': 1, '조사': 1, '착수': 1, '가한': 1, '노후화된': 1, '1968년산': 1, '헬기를': 1, '탔나': 1, '빈자리': 1, '클': 1, '분열된': 1, '실시...50일': 1, '안에': 1, '선출': 1, '빠진': 1, '열기로': 1, '벗,': 1, '기적': 1, '탈출': 1, '현지': 1, '관영': 1, '언론이': 1, '올린': 1, '사진': 1, '알고보니': 1, '200만': 1, '명이': 1, '속았다': 1, '북': 1, '소식': 1, '공화당은': 1, '책임은': 1, '이란에': 1, '극심한': 1, '적을': 1, '영향은?': 1, '교황,': 1, '어려운': 1, '시기': 1, '영적': 1, '친밀감': 1, '안갯속...중동': 1, '격랑,': 1, '중동정세는': 1, '對美': 1, '美와': 1, '핵협상-중동': 1, '불확실성': 1, '커질': 1, '검사시절': 1, '정치범': 1, '5000명': 1, '사형': 1, '하락': 1, '이란시민': 1, '호외': 1, '들고': 1, '망연자실': 1, '유력\\': 1, '급사에': 1, '소용돌이': 1, '최고지도자에': 1, '대한': 1, '비판': 1, '대신': 1, '받는': 1, '희생양': 1, '먹구름': 1, '초강경파': 1, '광장': 1, '외신사진': 1, '이슈人': 1, '故라이시': 1, '물결,': 1, '진영은': 1, '분노': 1, '반색': 1, '뽑는': 1, '율법전문회의,': 1, '신임': 1, '의장에': 1, '케르마니': 1, '당선': 1, '헬기부품': 1, '못구해': 1, '띄웠나': 1, '사망으로': 1, '이슬람': 1, '문화권': 1, '24시간': 1, '치르나': 1, '급부상한': 1, '암투': 1, '투쟁': 1, '통치': 1, '부활': 1, '헝클어진': 1, '대통령에': 1, '0.000025%': 1, '확률': 1, '항공사고로': 1, '목숨': 1, '지도자들': 1, '뭘': 1, '탔기에': 1, '돌연': 1, '내홍': 1, '시아파': 1, '성지': 1, '마슈하드에': 1, '외교1차관,': 1, '주한이란대사관': 1, '방문': 1, '조문': 1, '저항의': 1, '축': 1, '대거': 1, '미국에': 1, '죽음을': 1, '구호': 1, '인산인해': 1, '대통령실장': 1, '이륙': 1, '맑았다': 1, '하메네이의': 1, '포스트': 1, '선택은?': 1, '대통령이': 1, '탔던': 1, '탓에': 1, '추락했다?': 1, '국영방송': 1, '수백만': 1, '모여\\': 1, '횡설수설/조종엽': 1, '맞은': 1, '결론': 1, '힘': 1, '잃는': 1, '1차': 1, '보고서': 1, '1분30초': 1, '전까지': 1, '교신': 1, '수만명': 1, '확인했다': 1, '외무부': 1, '장관은': 1, '시계': 1, '현장,': 1, '구분도': 1, '어려워': 1, 'IAEA서': 1, '채택': 1, '못하도록': 1, '이번주': 1, '절차': 1, '개시...최대': 1, '난립': 1, '레이스': 1, '약': 1, '시론': 1, '사망과': 1, '고농축우라늄': 1, '늘려': 1, 'No.1': 1, '숨지자': 1, '공포': 1, '되살아났다': 1, '혹여': 1, '자극할까': 1, '두려운': 1, '바이든': 1, '신중': 1, '美-英佛은': 1, '바이든,': 1, '늘린': 1, '감싸기?...': 1, '결의안': 1, '추진': 1, '반대': 1, '이후': 1, '장악': 1, '새': 1, '국회': 1, '개회': 1, '연설서': 1, '건재': 1, '막으려': 1, '유럽': 1, '핵무기': 1, '3개': 1, '만들': 1, '있을': 1, '정도': 1, '폭탄급': 1, '커지는': 1, '비축': 1, '늘린다': 1, '피폐에': 1, '국민': 1, '불만은': 1, '고조': 1, '리더십과': 1, '이란-하마스': 1, '관계오늘,': 1, '박현도의': 1, '퍼스펙티브': 1, '악천후냐,': 1, '암살이냐': 1, '음모론에': 1, '세계는': 1, '지금': 1, '지도자는?': 1, '보궐대선': 1, '돌입': 1, '온건파는': 1, '심사': 1, '통과': 1, '어려울': 1, '유엔총회': 1, '행사': 1, '불참할': 1, '참모부': 1, '헬기사고,': 1, '영향': 1}
        wordcloud   = WordCloud(background_color='white',
                                font_path='/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf')
        wordcloud.generate_from_frequencies(self.m_word_counter)
        plt.figure(figsize=(10,10))
        plt.axis('off')
        plt.imshow(wordcloud)
        plt.show()            
    def TestReadXlsx(self, filename):
        print(f'# read xlsx file start.')
        df = pd.read_excel(filename)
        titles  = df['제목']
        print(titles)
        print(f'# read xlsx file end.')

def main(args):
    word_cloud_4_xlsx  = WordCloud4XLSX()
    word_cloud_4_xlsx.Run(args)
#    word_cloud_4_xlsx.TestReadXlsx('./data/Riasi_korean_news.xlsx')
#    word_cloud_4_xlsx.TestWordCloud()
    
if __name__ == '__main__':
    main(sys.argv)