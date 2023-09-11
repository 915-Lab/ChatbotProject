import json
import requests
import pandas as pd  # 엑셀 파일 읽기 위함
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill

"""
    대학백과 사이트 (https://www.univ100.kr)에서 동의대학교에 해당하는 모든 학과 Q&A 데이터 들고와서 학과별 엑셀 파일로 저장하기

    [분석]
    해당 사이트에서는 동적으로 데이터를 가지고 오고 있는 방식이다.
    그렇기 때문에 원본 데이터를 동적으로 들고오는 url을 network에서 가지고 오는 방식으로 데이터를 크롤링하는 방향으로 진행해야한다.

    학교 학과별 Q&A 데이터 요청 url : https://api.univ100.kr/find/qna/question/list?campusId=183&deptId=7853&keyword=&limit=20&offset=20

    campusId 쿼리 (대학교 번호) : 동의대학교 183
    deptId 쿼리 (학과 번호) : 창의소프트웨어공학부 7853
    
    대학백과 사이트에서는 Q&A 데이터를 20개씩 보여주고 있다.
    최초 요청할 때 offset=0, 다음 요청할 때 offset=20, 그 다음 요청할 때 offset=40 이런식으로 요청하는 방식
    
    [분석 결과]
    동의대학교만 사용한다 가정
    url에서 deptId를 학과 번호에 알맞게 변경해주고 offset만 계속해서 변경해주면 될 것 같다.


    [코딩 규칙]
        들여쓰기 형식 : 띄어쓰기 4번 (python 표준)
        스네이크 케이스 표기법 사용 : 파이썬은 PEP8을 통해 스네이크 케이스 방식의 네이밍 컨벤션을 권장함 ex) snake_case
"""


# 엑셀 파일의 값을 읽어서 학과, 학과번호 딕셔너리 변수 리스트를 반환해주는 함수
def dept_data_reader():
    file_path = 'data/univ100_deptId.xlsx'  # 학과 관련 정보 파일 경로
    df = pd.read_excel(file_path)

    file_data_list = []  # 각 행을 사전 딕셔너리 자료형으로 저장할 리스트

    # 파일의 각 행을 읽어서 학과, 학과번호 이렇게 딕셔너리 형태로 저장하기
    for index, row in df.iterrows():
        file_data = {}  # 각 행의 데이터를 저장할 딕셔너리 자료형

        file_data['학과'] = row['학과']
        file_data['학과번호'] = row['학과번호']

        file_data_list.append(file_data)

    return file_data_list


# 엑셀 파일로 저장하는 함수
def qna_save(dept_name, qna_list):
    print(qna_list)

    # 데이터프레임 생성
    df = pd.DataFrame(qna_list)

    # 엑셀 파일 생성
    excel_file_name = f'data/qna/{dept_name}.xlsx'
    df.to_excel(excel_file_name, sheet_name='QnA', index=False, header=True)



def data_crawler(dept_name, dept_id):
    """
    qna 데이터를 가져오는 함수
    반드시 header를 작성해줘야한다. 하지않을 경우 403 에러 발생
    header를 통해 로봇이  요청하는 것이 아닌 User-Agent를 지정해서 크롬 브라우저에서의 요청인것으로 인식하게 만들어줘야한다.

    :return:
    """

    os.makedirs('data/qna', exist_ok=True)

    dynamic_offset = 0  # 동적으로 변경되는 offset

    co = 1

    qna_list = []

    while True:
        base_url = f'https://api.univ100.kr/find/qna/question/list?campusId=183&deptId={dept_id}&keyword=&limit=20&offset={dynamic_offset}'
        qna = requests.get(base_url, headers={"User-Agent": "Mozilla/5.0"})  # headers 반드시 넣어줘야한다. 하지 않으면 403 에러 발생
        print(qna.status_code)  # 정상적으로 접속이 되었는지 확인

        qna.encoding = "utf-8"

        qna_data = json.loads(qna.text)  # json 형태 데이터 저장

        # json이 빈칸인 경우
        if len(qna_data['result']['questions']) == 0:
            break

        for q_data in qna_data['result']['questions']:
            print(f"{co}번 질문 : " + q_data['title'])
            qna_dict = {}
            qna_dict['질문'] = q_data['title']

            # 답변이 달려있지 않을 경우에는 answer가 없기 때문에 예외 처리
            if q_data.get('answer'):
                qna_dict['답변'] = q_data['answer']['text']
                print(f"{co}번 답변 : " + q_data['answer']['text'] + "\n")
            else:
                qna_dict['답변'] = '답변 없음'
                print(f"{co}번 답변 : 답변 없음\n")

            qna_list.append(qna_dict)

            co = co + 1

        dynamic_offset = dynamic_offset + 20

    return qna_list


if __name__ == '__main__':
    dept_data = dept_data_reader()

    for data in dept_data:
        qna = data_crawler(data['학과'], data['학과번호'])
        qna_save(data['학과'], qna)
