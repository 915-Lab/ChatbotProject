import json
import os

import pandas as pd
import requests


# 엑셀 파일로 저장하는 함수
def qna_save(qna_data, dept_name):
    print("--------------------------들어옴--------------------------")

    # 데이터프레임 생성
    df = pd.DataFrame(qna_data)

    for name in dept_name:
        # 학과별 데이터 필터링
        dept_qna_data = [data for data in qna_data if data.get('학과') == name]

        if dept_qna_data:
            # 데이터프레임 생성
            df = pd.DataFrame(dept_qna_data)

            # 엑셀 파일 이름 생성
            excel_file_name = f'data/qna/{name}.xlsx'

            # 엑셀 파일로 저장
            df.to_excel(excel_file_name, sheet_name='QnA', index=False, header=True)
            print(f"{name} 학과 엑셀 파일 저장 완료")



def data_crawler():
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
    dept_name_list_temp = []

    while True:
        base_url = f'https://api.univ100.kr/find/qna/question/list?campusId=183&keyword=&limit=20&offset={dynamic_offset}'
        qna = requests.get(base_url, headers={"User-Agent": "Mozilla/5.0"})  # headers 반드시 넣어줘야한다. 하지 않으면 403 에러 발생
        print(qna.status_code)  # 정상적으로 접속이 되었는지 확인

        qna.encoding = "utf-8"

        qna_data = json.loads(qna.text)  # json 형태 데이터 저장

        # json이 빈칸인 경우
        if len(qna_data['result']['questions']) == 0:
            break

        for q_data in qna_data['result']['questions']:

            # 카테고리가 입시 상담이고 질문하는 학과를 지정해둔 QnA 데이터만 저장
            if q_data['categoryName'] == '입시 상담' and len(q_data['deptName']) != 0:
                qna_dict = {}

                print(f"{co}번 질문 학과: " + q_data['deptName'])

                if '・' in q_data['deptName']:
                    temp = q_data['deptName'].replace('・', '_')
                    dept_name_list_temp.append(temp)
                    qna_dict['학과'] = temp
                elif '·' in q_data['deptName']:
                    temp = q_data['deptName'].replace('·', '_')
                    dept_name_list_temp.append(temp)
                    qna_dict['학과'] = temp
                else:
                    dept_name_list_temp.append(q_data['deptName'])
                    qna_dict['학과'] = q_data['deptName']

                print(f"{co}번 질문 : " + q_data['title'])
                qna_dict['질문'] = q_data['title']

                print(f"{co}번 질문 내용 : " + q_data['text'])
                qna_dict['내용'] = q_data['text']

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

    return qna_list, set(dept_name_list_temp)


if __name__ == '__main__':
    qna_data, dept_name = data_crawler()
    qna_save(qna_data, dept_name)
