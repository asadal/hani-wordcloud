import streamlit as st
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from konlpy.tag import Okt
from collections import Counter
from PIL import Image
import numpy as np
import openpyxl
import pandas as pd
from io import BytesIO
import os
import sys

# 환경 변수 설정 (Java 경로 지정)
# JAVA_HOME과 LD_LIBRARY_PATH를 설정하여 JVM을 찾을 수 있도록 합니다.
if 'JAVA_HOME' not in os.environ:
    # 일반적으로 OpenJDK 11의 경로
    os.environ['JAVA_HOME'] = '/usr/lib/jvm/java-11-openjdk-amd64'
    os.environ['LD_LIBRARY_PATH'] = os.environ.get('LD_LIBRARY_PATH', '') + ':/usr/lib/jvm/java-11-openjdk-amd64/lib/server'

# 페이지 설정
st.set_page_config(page_title="워드클라우드 생성기")

# 제목 및 설명
st.title("워드클라우드 생성기")
st.markdown("""
이 애플리케이션은 텍스트 파일, 엑셀 파일, 또는 직접 입력을 통해 단어 빈도수를 분석하고 워드클라우드를 생성합니다.
옵션을 설정하여 최대 단어 수, 테마, 마스크, 폰트를 선택할 수 있습니다.
""")

# 폰트 및 마스크 이미지 경로 설정
FONT_PATHS = {
    '나눔바른고딕': 'fonts/NanumBarunGothic.ttf',
    '나눔명조': 'fonts/NanumMyeongjo.ttf',
    'G마켓산스': 'fonts/GmarketSans.ttf',
    'Paperlogy': 'fonts/Paperlogy.ttf'
}

MASK_IMAGES = {
    '동그라미': 'masks/circle.jpg',
    '정사각형': 'masks/square.jpg',
    '세모': 'masks/triangle.jpg',
    '별': 'masks/star.jpg',
    '하트': 'masks/heart.jpg',
    '직사각형': 'masks/rectangle.jpg'
}

# 3단계 데이터 입력 선택
st.header("1. 데이터 입력")
input_method = st.radio("데이터 입력 방법을 선택하세요:", ("텍스트 파일 업로드", "엑셀 파일 업로드", "직접 입력"))

# 데이터프레임 초기화
df = pd.DataFrame(columns=['단어', '빈도수'])

if input_method == "텍스트 파일 업로드":
    txt_file = st.file_uploader("텍스트 파일을 업로드하세요(.txt)", type=["txt"])
    if txt_file is not None:
        try:
            text = txt_file.read().decode('utf-8')
            okt = Okt()
            words_with_pos = okt.pos(text, stem=True)
            desired_pos = {'Noun'}
            stopwords = set(['저희', '이제', '그리고', '하지만', '또한', '그', '그녀', '그들', '저'])
            tokens = [word for word, pos in words_with_pos if pos in desired_pos and word not in stopwords]
            tokens = [word for word in tokens if len(word) > 1]
            count = Counter(tokens)
            top_words = count.most_common(200)
            df = pd.DataFrame(top_words, columns=['단어', '빈도수'])
            st.success("텍스트 파일에서 단어를 성공적으로 추출했습니다.")
        except Exception as e:
            st.error(f"텍스트 파일 처리 중 오류가 발생했습니다: {e}")

elif input_method == "엑셀 파일 업로드":
    xlsx_file = st.file_uploader("엑셀 파일을 업로드하세요(.xlsx). 칼럼은 '단어'와 '빈도수'로 설정해주세요.", type=["xlsx"])
    if xlsx_file is not None:
        try:
            wb = openpyxl.load_workbook(xlsx_file, data_only=True)
            sheets = wb.sheetnames
            selected_sheet = st.selectbox("데이터를 가져올 시트를 선택하세요:", sheets)
            sheet = wb[selected_sheet]
            data = sheet.values
            cols = next(data)
            df = pd.DataFrame(data, columns=cols)
            if '단어' in df.columns and '빈도수' in df.columns:
                df = df[['단어', '빈도수']].dropna()
                df = df.sort_values(by='빈도수', ascending=False).reset_index(drop=True)
                st.success("엑셀 파일에서 단어를 성공적으로 불러왔습니다.")
            else:
                st.error("엑셀 파일에 '단어'와 '빈도수' 컬럼이 존재하지 않습니다.")
        except Exception as e:
            st.error(f"엑셀 파일 처리 중 오류가 발생했습니다: {e}")

elif input_method == "직접 입력":
    st.subheader("단어와 빈도수를 직접 입력하세요 (각 단어와 빈도수를 콤마로 구분하고, 줄 바꿈으로 구분)")
    user_input = st.text_area("입력 예시:\n사람,536\n사랑,423\n행복,389")
    if st.button("단어 추출"):
        try:
            lines = user_input.strip().split('\n')
            words = []
            frequencies = []
            for line in lines:
                if ',' in line:
                    word, freq = line.split(',', 1)
                    word = word.strip()
                    try:
                        freq = int(freq.strip())
                        words.append(word)
                        frequencies.append(freq)
                    except ValueError:
                        st.warning(f"빈도수가 유효하지 않은 단어를 건너뜁니다: {line}")
            df = pd.DataFrame({'단어': words, '빈도수': frequencies})
            df = df.sort_values(by='빈도수', ascending=False).reset_index(drop=True)
            st.success("직접 입력한 단어를 성공적으로 추출했습니다.")
        except Exception as e:
            st.error(f"직접 입력 처리 중 오류가 발생했습니다: {e}")

# 데이터프레임 표시 및 불필요한 줄 삭제
if not df.empty:
    st.header("2. 데이터 검토 및 정제")
    st.write("데이터프레임을 확인하고, 불필요한 단어를 삭제하세요.")
    selected_words = st.multiselect("삭제할 단어를 선택하세요:", df['단어'].tolist())
    if selected_words:
        df = df[~df['단어'].isin(selected_words)]
        st.success(f"{len(selected_words)}개의 단어를 삭제했습니다.")
    st.dataframe(df)
else:
    st.warning("데이터가 없습니다. 데이터를 입력하세요.")

# 옵션 설정
st.header("3. 워드클라우드 옵션 설정")
col1, col2, col3, col4 = st.columns(4)

with col1:
    max_words = st.slider("최대 단어 수", min_value=20, max_value=200, value=100, step=10)

with col2:
    theme = st.selectbox("테마 선택 (Colormap)", options=['viridis', 'plasma', 'inferno', 'magma', 'cividis'])

with col3:
    mask_choice = st.selectbox("마스크 선택", options=list(MASK_IMAGES.keys()))

with col4:
    font_choice = st.selectbox("폰트 선택", options=list(FONT_PATHS.keys()))

st.write("JAVA_HOME:", os.environ.get('JAVA_HOME'))
st.write("LD_LIBRARY_PATH:", os.environ.get('LD_LIBRARY_PATH'))
st.write("Java 버전:")
os.system("java -version")

# 워드클라우드 생성 버튼
if st.button("워드클라우드 생성"):
    if not df.empty:
        try:
            # 선택한 마스크 이미지 로드
            mask_path = MASK_IMAGES[mask_choice]
            mask_image = Image.open(mask_path).convert("L")
            mask_array = np.array(mask_image)
            
            # 선택한 폰트 경로
            font_path = FONT_PATHS[font_choice]
            if not os.path.exists(font_path):
                st.error(f"폰트 파일이 존재하지 않습니다: {font_path}")
                st.stop()
            
            # 단어 빈도수 딕셔너리로 변환
            word_freq = dict(zip(df['단어'], df['빈도수']))
            
            # 워드클라우드 생성
            wc = WordCloud(
                font_path=font_path,
                background_color="white",
                mask=mask_array,
                width=1800,
                height=1800,
                scale=2.0,
                max_words=max_words,
                colormap=theme,
                random_state=42,
                min_font_size=5,
                relative_scaling=0.5,
                prefer_horizontal=0.9,
                collocations=False
            )
            
            wc.generate_from_frequencies(word_freq)
            
            # 워드클라우드 이미지 생성
            fig, ax = plt.subplots(figsize=(10,10))
            ax.imshow(wc, interpolation="bilinear")
            ax.axis("off")
            plt.tight_layout(pad=0)
            st.pyplot(fig)
            
            # 이미지 다운로드
            img_buffer = BytesIO()
            wc.to_file(img_buffer, format='PNG')
            img_buffer.seek(0)
            st.download_button(
                label="워드클라우드 이미지 다운로드",
                data=img_buffer,
                file_name="wordcloud.png",
                mime="image/png"
            )
            
            # 데이터 다운로드
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Word Frequency')
            excel_buffer.seek(0)
            st.download_button(
                label="데이터 엑셀 파일 다운로드",
                data=excel_buffer,
                file_name="word_count_top_100.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"워드클라우드 생성 중 오류가 발생했습니다: {e}")
    else:
        st.error("데이터프레임이 비어있습니다. 데이터를 입력하고 불필요한 단어를 삭제하세요.")

# Reload 버튼
if st.button("Reload"):
    st.rerun()
