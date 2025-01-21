import os
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

# 환경 변수 설정 (Java 경로 지정)
# Streamlit Cloud에서는 기본적으로 OpenJDK 11이 설치되어 있으므로, JAVA_HOME과 LD_LIBRARY_PATH를 설정합니다.
os.environ['JAVA_HOME'] = '/usr/lib/jvm/java-11-openjdk-amd64'
os.environ['LD_LIBRARY_PATH'] = '/usr/lib/jvm/java-11-openjdk-amd64/lib/server'

# 페이지 설정
st.set_page_config(page_title="WordCloud Generator")

# 제목 및 Reload 버튼을 옆에 배치하기 위한 컬럼 설정
col_title, col_reload = st.columns([4, 1])

with col_title:
    st.title("WordCloud Generator")

with col_reload:
    if st.button("Reload ⟳"):
        # 세션 상태 초기화
        st.session_state.wordcloud_generated = False
        st.session_state.img_bytes = None
        st.session_state.excel_bytes = None
        st.session_state.custom_mask = None
        # 애플리케이션 재실행
        st.rerun()  # 최신 Streamlit에서는 st.rerun() 사용 가능. 지원되지 않으면 st.experimental_rerun() 사용

st.markdown("""
텍스트 파일, 엑셀 파일, 또는 직접 입력을 통해 단어 빈도수를 분석하고 워드클라우드를 생성합니다.
최대 단어 수, 테마, 마스크, 폰트를 선택할 수 있습니다.
""")

# 폰트 및 마스크 이미지 경로 설정
FONT_PATHS = {
    '나눔바른고딕': 'fonts/NanumBarunGothic.ttf',
    '나눔명조': 'fonts/NanumMyeongjo.ttf',
    'G마켓산스': 'fonts/GmarketSans.ttf',
    '페이퍼로지': 'fonts/Paperlogy.ttf'
}

MASK_IMAGES = {
    '▆': 'masks/rectangle.jpg',
    '■': 'masks/square.jpg',
    '●': 'masks/circle.jpg',
    '★': 'masks/star.jpg',
    '▲': 'masks/triangle.jpg',
    '♥': 'masks/heart.jpg',
    '이미지 업로드': '이미지 업로드'
}

# 초기 세션 상태 설정
if 'wordcloud_generated' not in st.session_state:
    st.session_state.wordcloud_generated = False

if 'img_bytes' not in st.session_state:
    st.session_state.img_bytes = None

if 'excel_bytes' not in st.session_state:
    st.session_state.excel_bytes = None

if 'custom_mask' not in st.session_state:
    st.session_state.custom_mask = None

# 1. 데이터 입력 섹션
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

# 2. 데이터 검토 및 정제 섹션
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

# 3. 워드클라우드 옵션 설정 섹션
st.header("3. 워드클라우드 옵션 설정")
col1, col2, col3, col4 = st.columns(4)

with col1:
    max_words = st.slider("최대 단어 수", min_value=20, max_value=200, value=100, step=10)

with col2:
    theme = st.selectbox("테마 선택 (Colormap)", options=['viridis', 'plasma', 'inferno', 'magma', 'cividis'])

with col3:
    mask_choice = st.selectbox("마스크 선택", options=list(MASK_IMAGES.keys()))
    
    # '직접 업로드' 선택 시 파일 업로더 표시
    if mask_choice == '이미지 업로드':
        custom_mask_file = st.file_uploader("업로드할 마스크 이미지를 선택하세요(jpg)", type=["jpg", "jpeg"])
        if custom_mask_file is not None:
            try:
                custom_mask_image = Image.open(custom_mask_file).convert("L")  # 그레이스케일로 변환
                custom_mask_array = np.array(custom_mask_image)
                st.session_state.custom_mask = custom_mask_array
                st.success("마스크 이미지를 성공적으로 업로드했습니다.")
            except Exception as e:
                st.error(f"마스크 이미지 처리 중 오류가 발생했습니다: {e}")
    else:
        st.session_state.custom_mask = None  # 다른 마스크 선택 시 커스텀 마스크 초기화

with col4:
    font_choice = st.selectbox("폰트 선택", options=list(FONT_PATHS.keys()))

# 워드클라우드 생성 버튼
if st.button("워드클라우드 생성"):
    if not df.empty:
        try:
            # 선택한 마스크 이미지 로드
            if mask_choice == '이미지 업로드':
                if st.session_state.custom_mask is not None:
                    mask_array = st.session_state.custom_mask
                else:
                    st.error("마스크 이미지를 업로드하지 않았습니다.")
                    st.stop()
            else:
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
                width=3600,
                height=3600,
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
            fig, ax = plt.subplots(figsize=(10, 10))
            ax.imshow(wc, interpolation="bilinear")
            ax.axis("off")
            plt.tight_layout(pad=0)
            st.pyplot(fig)
        
            # 워드클라우드 이미지 버퍼에 저장
            img_buffer = BytesIO()
            image = wc.to_image()
            image.save(img_buffer, format='PNG')
            img_bytes = img_buffer.getvalue()
            st.session_state.img_bytes = img_bytes
        
            # 데이터 엑셀 파일 버퍼에 저장
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Word Frequency')
            excel_bytes = excel_buffer.getvalue()
            st.session_state.excel_bytes = excel_bytes
        
            # 세션 상태 업데이트
            st.session_state.wordcloud_generated = True
        
            st.success("워드클라우드를 성공적으로 생성했습니다.")
        
        except Exception as e:
            st.error(f"워드클라우드 생성 중 오류가 발생했습니다: {e}")
    else:
        st.error("데이터프레임이 비어 있습니다. 데이터를 입력하고 불필요한 단어를 삭제하세요.")

# 다운로드 버튼을 세션 상태를 통해 지속적으로 표시
if st.session_state.get('wordcloud_generated', False):
    st.header("다운로드")
    # 워드클라우드 이미지 다운로드 버튼
    if st.session_state.img_bytes:
        st.download_button(
            label="워드클라우드 이미지 다운로드",
            data=st.session_state.img_bytes,
            file_name="wordcloud.png",
            mime="image/png"
        )

    # 데이터 엑셀 파일 다운로드 버튼
    if st.session_state.excel_bytes:
        st.download_button(
            label="데이터 파일 다운로드",
            data=st.session_state.excel_bytes,
            file_name="word_count_top_100.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
