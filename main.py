import pandas as pd
from collections import Counter
import streamlit as st
import matplotlib.pyplot as plt
from wordcloud import WordCloud
import re
import io

st.title('설문조사')
st.set_option('deprecation.showPyplotGlobalUse', False)

st.info(" 엑셀 파일 업로드", icon="💾")

# 엑셀 파일 업로드
st.subheader("엑셀 평균 계산")
uploaded_file = st.file_uploader("평균 - 설문조사파일 업로드", type=['xlsx'])

if uploaded_file is not None:
    # 엑셀 파일 읽기
    df = pd.read_excel(uploaded_file)

    # 문자열 열 제거 (1열은 문자열이라고 가정)
    numeric_df = df.select_dtypes(include='number')

    # 열마다 평균과 총 셀 개수, 빈 셀 개수 계산
    column_means = numeric_df.mean(skipna=True).round(2)  # 소수점 둘째 자리까지 반올림
    total_cell_counts = numeric_df.count()  # 총 셀 개수
    empty_cell_counts = numeric_df.isnull().sum()  # 빈 셀 개수

    # 결과 출력
    st.write("숫자 열의 평균, 응답, 미응답:")
    result_df = pd.DataFrame({'평균': column_means, '응답': total_cell_counts, '미응답': empty_cell_counts})
    st.write(result_df)

st.subheader("과락 확인")
score = st.radio("수업 유형(파이썬 기초, 전처리)", ["SELECT", "(PYTHON)", "(PANDAS)"])

if score == 'SELECT':
    st.write("------------유형 선택------------")

# Streamlit에서 파일 업로드
elif score == "(PYTHON)":
    uploaded_file = st.file_uploader("대시보드 파일 업로드")
    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file)

        # 사용자가 지정한 위치에 열을 불러오는 방식
        positions = [7, 9, 11]  # 7열, 9열, 11열에 해당하는 위치

        # 사용자가 지정한 시트 읽기
        sheet_name1 = '_응시 필수_ 파이썬 과제(3개 문제 중 1개만 진행)'
        df1 = pd.read_excel(xls, sheet_name1)

        # 결과를 저장할 빈 데이터프레임을 생성합니다.
        result_df = pd.DataFrame(columns=['이름', '이메일', '과제 최대값'])

        # 2행부터 끝 행까지 각 행의 최대값을 찾는다.
        for i in range(1, len(df1)):
            max_val = df1.iloc[i, positions].max()
            # 최대값이 80 미만인 경우에만 출력한다.
            if max_val < 80:
                result_df.loc[len(result_df)] = [df1.iloc[i, 1], df1.iloc[i, 2], max_val]

        # 결과를 표 형식으로 출력합니다.
        st.table(result_df)

        # 출력된 값의 개수를 카운트합니다.
        st.write(f"과제 과락 인원: {len(result_df)}")
        st.write("----------------------------------------------------------------------------------")

#---------------------------------------------------------------------------------------

        sheet_name2 = '_응시 필수_ 사후 테스트'

        df2 = pd.read_excel(xls, sheet_name2)

        df2 = df2.sort_values(df2.columns[6])

        filtered_df2 = df2[df2[df2.columns[6]] < 80]

        total_cells2 = df2[df2.columns[6]].count()
        less_than_80_cells2 = filtered_df2[df2.columns[6]].count()

        position_test = [1,2,6]

        st.write(f"테스트 응시 인원: {total_cells2} /// 80 미만 인원: {less_than_80_cells2}")
        st.write(filtered_df2.iloc[:,position_test])

# Streamlit에서 파일 업로드
else:
    uploaded_file = st.file_uploader("대시보드 파일 업로드")
    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file)

        # 사용자가 지정한 위치에 열을 불러오는 방식
        positions = [7, 9, 11]  # 7열, 9열, 11열에 해당하는 위치

        # 사용자가 지정한 시트 읽기
        sheet_name1 = '_응시 필수_ 실습 과제 (3개 문제 중 1개만 진행)'
        df1 = pd.read_excel(xls, sheet_name1)

        # 결과를 저장할 빈 데이터프레임을 생성합니다.
        result_df = pd.DataFrame(columns=['이름', '이메일', '과제 최대값'])

        # 2행부터 끝 행까지 각 행의 최대값을 찾는다.
        for i in range(1, len(df1)):
            max_val = df1.iloc[i, positions].max()
            # 최대값이 80 미만인 경우에만 출력한다.
            if max_val < 80:
                result_df.loc[len(result_df)] = [df1.iloc[i, 1], df1.iloc[i, 2], max_val]

        # 결과를 표 형식으로 출력합니다.
        st.table(result_df)

        # 출력된 값의 개수를 카운트합니다.
        st.write(f"과제 과락 인원: {len(result_df)}")
        st.write("----------------------------------------------------------------------------------")

#-------------------------------------------------------------------------------------
        sheet_name2 = '_응시 필수_ 사후 테스트'

        df2 = pd.read_excel(xls, sheet_name2)

        df2 = df2.sort_values(df2.columns[6])

        filtered_df2 = df2[df2[df2.columns[6]] < 80]

        total_cells2 = df2[df2.columns[6]].count()
        less_than_80_cells2 = filtered_df2[df2.columns[6]].count()

        position_test = [1,2,6]

        st.write(f"테스트 응시 인원: {total_cells2} /// 80 미만 인원: {less_than_80_cells2}")
        st.write(filtered_df2.iloc[:,position_test])

st.write("--------------------------------------------------------------------------------")
    
# 엑셀 파일 업로드
st.subheader("텍스트 분석")
uploaded_file = st.file_uploader("텍스트 - 파일 업로드", type=['xlsx'])

# 엑셀 파일에서 텍스트 분석하는 함수 정의
def analyze_text_in_excel(df, column):
    # 선택한 열에서 텍스트 추출
    text_data = df[column]

    # 추출한 텍스트를 단어로 나누기 위해 모든 텍스트를 하나의 문자열로 합침
    all_text = ' '.join(text_data.dropna().astype(str).values)

    # 단어 단위로 나누기
    words = re.findall(r'\b\w+\b', all_text.lower())  # 소문자로 변환 후 단어 추출

    return words

# 엑셀 파일에서 숫자 열의 평균을 계산하는 함수 정의
def calculate_column_mean(df, column):
    return round(df[column].mean(),2)

# 불용어를 제거하는 함수 정의
def remove_stopwords(words, stopwords):
    return [word for word in words if word not in stopwords]

# wordcloud를 생성하여 시각화하는 함수 정의
def generate_wordcloud(words):
    # 한글 폰트 경로 지정
    font_path = 'C:\\Users\\21-01-131\\AppData\\Local\\Microsoft\\Windows\\Fonts\\NanumGothicBold.ttf'

    # 단어 빈도로부터 WordCloud 생성
    wordcloud = WordCloud(width=800, height=400, font_path=font_path, background_color='white').generate(' '.join(words))

    # 시각화

    plt.figure(figsize=(10, 5))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    st.pyplot()

if uploaded_file is not None:
    # 엑셀 파일 읽기
    df = pd.read_excel(uploaded_file)

    # 열 선택
    selected_columns = st.multiselect('분석할 열 선택', df.columns)

    for selected_column in selected_columns:
        if df[selected_column].dtype == 'object':  # 텍스트 열인 경우
            # 불용어 설정
            stopwords = ['을', '를', '가', 'nan', '수', '있는', '있어서', '대한', '점이', '들', '더', '것','좀','너무','조금', 
             '있습니다', '있으면', '좀더','하는','할','다소','여러','합니다','하는','등','보니','한','대해','없는'
             ,'될', '볼','잘','보임','안','거','위한','해','건','같음','좋겠다','있도록','함께','정도','경우','오히려','등에서'
             '빠른','있었는데','것이','때가','모두','필요한','않아','갈수록','되지','때','아주','다른','일부','간혹','전혀','하나'
             ,'같다','해당','정도면','매우','위하여','했으면','완전','하기에는','다수','하다','동안','것을','느낌은','좋겠고'
             ,'외','위해','않도록','됐으면', '나오는데','아니면','있게','제','또는','자꾸','같네요','아니라','개','있고'
             '다만', '그부분은', '한번씩','있다는','하에','되는','알고있는','않아서','될때','더많은','많다', '통해'
             ,'등에서', '비해', '내야지', '있어야', '있고','다만','같은','같기도','알','있어','있었다','주셔서','많은'
             '알게','제일','된다면','하고','맞게','우선','대해서','둘','굉장히','받을수','따로','아직','많지만','1','2','없이'
             ,'있었던','까지','않고','수','같습니다','점','주시는','있었음','될거','될','거','알기엔','틈이','틈','아무리','점점','점'
             ,'됨','뜬것','싶었던','싶은','첫','그','저',"경우도","경우","수도","시간이","아무래도","다음에는"
             ,"정말","덕분에","고려","찾아","되고","떠서","있을","시","일찍","다시","하기는","이런","진행이","그리고"
             ,"무엇보다","접하는","즉시","및","내에서","내","잡는게","약간","하여","실제","마다","위로","있었습니다"
             ,"해주셔서","하지만","주었으면","어디까지","해주시는지","달리다보면", "되어","점과","해볼수","알기","않은","해보는","해주셨고"
             ,"들을","다를","되어","vs","전까지는","되었습니다","이른","통한","알려주어서","아실거라고","아래로","낼려고","있다고","다"
             ,"알려주면","판단됨니다","듯","먼저","보기"]  # 불용어 리스트에 포함시킬 단어들을 지정

            # 엑셀 파일에서 텍스트 추출 및 분석
            words = analyze_text_in_excel(df, selected_column)
            filtered_words = remove_stopwords(words, stopwords)

            # 단어 빈도 계산
            word_counts = Counter(filtered_words)

            # wordcloud 생성 및 시각화
            st.markdown(f"{selected_column}")
            generate_wordcloud(filtered_words)
        else:  # 숫자 열인 경우
            # 선택한 열의 평균 계산
            column_mean = calculate_column_mean(df, selected_column)

            # 결과 출력
            st.markdown(f"### ({selected_column}  | 평균: {column_mean})")

st.write("--------------------------------------------------------------------------------")
st.subheader("사용 X")
survey = st.radio("설문조사 유형(사전, 사후)", ["SELECT", "(BEFORE)", "(AFTER)"])

# 엑셀 파일 업로드

if survey == 'SELECT':
    st.write("------------유형 선택------------")

elif survey == '(BEFORE)':
    st.subheader("사전설문조사 - 사용 NO")
    uploaded_file = st.file_uploader("사전 - 파일 업로드", type=['xlsx'])
    if uploaded_file is not None:
        # 엑셀 파일 읽기
        df = pd.read_excel(uploaded_file)

        # 고정값으로 범위 지정
        start_row, end_row = 0, len(df)-1
        start_col, end_col = 6, len(df.columns)-1

        # 지정한 범위에서 빈 셀 개수 계산
        selected_range = df.iloc[start_row:end_row+1, start_col:end_col+1]
        empty_cells_per_row = selected_range.isnull().sum(axis=1)

        # 사용자가 지정한 숫자를 초과하는 빈 셀 개수를 가진 행 삭제
        st.warning("설문조사 미참여 인원 삭제 임계값은 아래 해당 값을 넣어주세요")

        spotfire = 3
        python = 4
        data = 5

        st.write(f"**spotfire 기초 및 생성형AI, Skillup-AI = {spotfire}**")
        st.write(f"**spotfire 심화 및 (C/D)파이썬 기초, Skillup-데이터분석, Skillup-Spotfire = {python}**")
        st.write(f"**(C/D)전처리 = {data}**")

        threshold = st.number_input("설문조사 미참여 인원 삭제 임계값:", value=4)
        rows_to_remove = empty_cells_per_row[empty_cells_per_row > threshold].index
        df.drop(rows_to_remove, inplace=True)

        # 각 행별로 가장 많이 등장한 정수값 찾기
        most_common_values = df.iloc[start_row:end_row+1, start_col:end_col+1].mode(axis=1)

        # 사용자가 지정한 값이 아닌 가장 많이 등장한 정수값으로 빈 셀 채우기
        for i, row_index in enumerate(most_common_values.index):
            most_common_value = most_common_values.loc[row_index].iloc[0]
            for col_name in df.columns[start_col:end_col+1]:
                if pd.isna(df.loc[row_index, col_name]):
                    df.loc[row_index, col_name] = most_common_value


        # 새로운 엑셀 파일로 저장
        output_filename = st.text_input("저장할 파일명을 입력해주세요(ex.수업명(사전)_강사이름_날짜):", value="새로운_파일_사전.xlsx")
        if st.button("저장하기"):
            # 새로운 엑셀 파일을 메모리에 저장
            towrite = io.BytesIO()
            df.to_excel(towrite, index=False, engine='xlsxwriter')
            towrite.seek(0)

            # 다운로드 버튼 생성
            st.download_button(
                label="Download Excel File",
                data=towrite,
                file_name=output_filename,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            st.success(f"새로운 엑셀 파일 '{output_filename}'로 저장되었습니다.")


else:
    # 엑셀 파일 업로드
    st.subheader("사후설문조사 - 사용 NO")
    uploaded_file = st.file_uploader("사후 -  파일 업로드", type=['xlsx'])

    if uploaded_file is not None:
        # 엑셀 파일 읽기
        df = pd.read_excel(uploaded_file)

        # 1. 사용자가 범위를 지정
        start_row, end_row = 0, 200
        start_col, end_col = 6, 30

        # 1-1. 
        start_row1, end_row1 = 0, 200
        start_col1, end_col1 = 6, 15

        # 1-2
        start_row2, end_row2 = 0, 200
        start_col2, end_col2 = 18, 20

        # 1-3
        start_row3, end_row3 = 0, 200
        start_col3, end_col3 = 21, 30


        # 2. 빈 셀이 많은 행 제거
        spotfire1 = 18
        python1 = 19
        data1 = 20

        st.write(f"**spotfire 기초 및 생성형AI = {spotfire1}**")
        st.write(f"**spotfire 심화 및 (C/D)파이썬 기초 = {python1}**")
        st.write(f"**(C/D)전처리 = {data1}**")

        threshold = st.number_input("설문조사 미참여 인원 삭제 임계값", value=18)
        empty_cells_per_row = df.iloc[start_row:end_row+1, start_col:end_col+1].isnull().sum(axis=1)
        rows_to_remove = empty_cells_per_row[empty_cells_per_row > threshold].index
        df.drop(rows_to_remove, inplace=True)

        # 문자열 열 제거 (1열은 문자열이라고 가정)
        numeric_df = df.select_dtypes(include='number')

        # 3. 빈 셀 채우기
        fill_value1 = st.number_input("1번부터 10번까지 채우기 값", value=5)
        fill_value2 = st.number_input("13번부터 15번까지 채우기 값", value=5)
        fill_value3 = st.number_input("16번부터 끝까지 채우기 값", value=4)

        # 열마다 평균과 총 셀 개수, 빈 셀 개수 계산
        column_means = numeric_df.mean(skipna=True).round(2)  # 소수점 둘째 자리까지 반올림
        total_cell_counts = numeric_df.count()  # 총 셀 개수
        empty_cell_counts = numeric_df.isnull().sum()  # 빈 셀 개수

        # 빈 셀 채우기
        df.iloc[start_row1:end_row1+1, start_col1:end_col1+1] = df.iloc[start_row1:end_row1+1, start_col1:end_col1+1].fillna(fill_value1)
        df.iloc[start_row2:end_row2+1, start_col2:end_col2+1] = df.iloc[start_row2:end_row2+1, start_col2:end_col2+1].fillna(fill_value2)
        df.iloc[start_row3:end_row3+1, start_col3:end_col3+1] = df.iloc[start_row3:end_row3+1, start_col3:end_col3+1].fillna(fill_value3)

        # 열마다 채운 값으로 평균 다시 계산
        column_means_updated = df.select_dtypes(include='number').mean(skipna=True).round(2)  # 소수점 둘째 자리까지 반올림

        # 결과 출력
        st.write("숫자 열의 평균, 총 셀 개수, 빈 셀 개수:")
        result_df = pd.DataFrame({'평균': column_means, '계산된 셀 개수': total_cell_counts, '빈 셀 개수': empty_cell_counts})
        st.write(result_df)

        st.write("빈 셀을 채운 후 숫자 열의 평균:")
        st.write(column_means_updated)

        # 새로운 엑셀 파일로 저장
        output_filename = st.text_input("저장할 파일명을 입력해주세요(ex.수업명(사전)_강사이름_날짜):", value="새로운_파일_사전.xlsx")
        if st.button("저장하기"):
            # 새로운 엑셀 파일을 메모리에 저장
            towrite = io.BytesIO()
            df.to_excel(towrite, index=False, engine='xlsxwriter')
            towrite.seek(0)

            # 다운로드 버튼 생성
            st.download_button(
                label="Download Excel File",
                data=towrite,
                file_name=output_filename,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            st.success(f"새로운 엑셀 파일 '{output_filename}'로 저장되었습니다.")