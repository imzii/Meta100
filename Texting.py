import os
import datetime
import requests
from bs4 import BeautifulSoup


# 2. 등급 계산 함수
def calculate_grade(score):
    if score >= 90:
        return 1
    elif score >= 80:
        return 2
    elif score >= 70:
        return 3
    elif score >= 60:
        return 4
    elif score >= 50:
        return 5
    elif score >= 40:
        return 6
    elif score >= 30:
        return 7
    elif score >= 20:
        return 8
    else:
        return 9


# 3. 틀린 문제 유형 표시 함수
def problem_type(errors):
    type_dict = {
        "듣기": list(range(1, 18)),
        "글의 목적": [18],
        "글의 분위기": [19],
        "대의 파악": [20, 22, 23, 24],
        "함의 추론": [21],
        "도표 이해": [25],
        "내용 일치 / 불일치": [26],
        "실용문 일치": [27, 28],
        "어법성 판단": [29],
        "단어 쓰임 판단": [30],
        "빈칸 추론": list(range(31, 35)),
        "무관한 문장 고르기": [35],
        "문단 순서 맞히기": [36, 37],
        "주어진 문장 넣기": [38, 39],
        "요약문 완성": [40],
        "기본 장문 독해": [41, 42],
        "복합 문단 독해": list(range(43, 46))
    }

    result = []
    for t, nums in type_dict.items():
        temp = [str(n) for n in errors if n in nums]
        if temp:
            result.append(", ".join(temp) + f"({t})")

    return ", ".join(result)


def get_suneung_date():
    url = "https://www.google.com/search?q=%EC%88%98%EB%8A%A5+%EB%82%A0%EC%A7%9C"
    response = requests.get(url)

    # HTML을 BeautifulSoup로 파싱
    soup = BeautifulSoup(response.text, 'html.parser')

    # selector로 원하는 요소 선택
    selector = '#EtGB6d > div > div > span > div > div:nth-child(2) > div.zCubwf > div:nth-child(1)'
    element = soup.select_one(selector)

    # 요소의 텍스트 가져오기
    if element:
        text = element.get_text()
    else:
        text = '2023년 11월 16일'
    return text


# 4. 문자 생성 함수
def generate_message(year, month, grade, student_name, student_score, wrong_problem_numbers, suneung_date):
    if month == 11 and grade == 3:
        test_category = f'{year}학년도 수능'
    else:
        test_category = f'고등학교 {grade}학년 {year}학년도 {month}월'
    message = f"""
안녕하세요.
메타백영어학원입니다.
{student_name} 학생의 모의고사 수업에서 실시한 {test_category} 모의고사 결과 알려드립니다.
등급: {calculate_grade(student_score)}등급
총점: {student_score}점
틀린 문제: {problem_type(wrong_problem_numbers)}

{suneung_date} 입시의 마지막 수능까지 최선을 다해 달려가겠습니다.
감사합니다.
"""
    return message


# 5. 파일 이름 결정 함수
def determine_filename(year, month, grade):
    counter = 1
    filename = f"{year}년 {month}월 고{grade} 모의고사 문자 {counter}차.txt"
    while os.path.exists(filename):
        counter += 1
        filename = f"{year}년 {month}월 고{grade} 모의고사 문자 {counter}차.txt"
    return filename


# 6. 파일 저장 함수
def save_to_file(filename, content):
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(content)


def input_year(current_year):
    while True:
        try:
            year = int(input("모의고사 년도를 입력하세요: "))
            if 2003 <= year <= (current_year + 1):
                return year
            else:
                print(f"2003에서 {current_year + 1} 사이의 년도를 입력하세요.")
        except ValueError:
            print("올바른 형식의 숫자를 입력하세요. 다시 시도하세요.")


def input_month():
    while True:
        try:
            month = int(input("모의고사 시행 월을 입력하세요: "))
            if 1 <= month <= 12:
                return month
            else:
                print("1월부터 12월 사이의 숫자를 입력하세요.")
        except ValueError:
            print("올바른 형식의 숫자를 입력하세요. 다시 시도하세요.")


def input_grade():
    while True:
        try:
            grade = int(input("모의고사 대상 학년을 입력하세요: "))
            if 1 <= grade <= 3:
                return grade
            else:
                print("1학년부터 3학년 사이의 숫자를 입력하세요.")
        except ValueError:
            print("올바른 형식의 숫자를 입력하세요. 다시 시도하세요.")


def input_student_score():
    while True:
        try:
            student_score = int(input("학생의 점수를 입력하세요: "))
            if 0 <= student_score <= 100:
                return student_score
            else:
                print("0점부터 100점 사이의 숫자를 입력하세요.")
        except ValueError:
            print("올바른 형식의 점수를 입력하세요. 다시 시도하세요.")


def input_wrong_problem_numbers():
    while True:
        wrong_problem_numbers = input("학생이 틀린 문제 번호를 ','로 구분하여 입력하세요: ").split(',')
        try:
            wrong_problem_numbers = [int(item.strip()) for item in wrong_problem_numbers]
            return wrong_problem_numbers
        except ValueError:
            print("올바른 형식의 숫자를 입력하세요. 다시 시도하세요.")


def main():
    current_year = datetime.datetime.now().year
    year = input_year(current_year)
    month = input_month()
    grade = input_grade()
    suneung_date = get_suneung_date()

    all_messages = []

    while True:
        student_name = input("학생의 이름을 입력하세요 (입력 종료를 원하시면 '종료'를 입력): ")
        if student_name == '종료':
            break
        student_score = input_student_score()
        wrong_problem_numbers = input_wrong_problem_numbers()

        message = generate_message(year, month, grade, student_name, student_score, wrong_problem_numbers, suneung_date)
        all_messages.append(message)

    content = "\n\n".join(all_messages)
    filename = determine_filename(year, month, grade)
    save_to_file(filename, content)
    print(f"'{filename}' 파일로 저장되었습니다.")


if __name__ == '__main__':
    main()
