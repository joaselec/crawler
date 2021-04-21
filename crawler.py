import requests
import urllib
from bs4 import BeautifulSoup
from openpyxl import Workbook
import logging


# 로거 생성
logger = logging.getLogger()
# 로그 출력 레벨
logger.setLevel(logging.INFO)
# 로그 출력 형식
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
# 로그 출력
stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)

# log를 파일에 출력
file_handler = logging.FileHandler("clawler.log")
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)


# 엑셀 객체 생성
wb = Workbook()
# 시트 생성
ws = wb.create_sheet("대출나라")
# 입력 활성화
ws = wb.active

for i in range(21, 1091):
    # for i in range(1092, 3000):
    try:
        # 1.url 선언 (count ++ until 10000)
        url = "http://xn--vk1bw4xgkai32b.com/inc/MebCorpTel_Pop.php?u_idx="
        url = url + str(i)
        logger.info("번호:" + str(i))
        # print("실행 번호:" + str(i))
        # print(url)

        # 2.크롤링 및 파싱

        # requst로 html 가져오기
        html = requests.get(url).text
        # print(response.content)

        # bs 객체로 변환
        # 파싱 -> bs 객체는 fild, fild_all 함수 제공
        bs = BeautifulSoup(html, "html.parser")

        # 테이블을 가져온다
        table = bs.find("table")
        # print(table)

        # 테이블에 정보가 없다면 continue
        if table.find_all("tr")[0].find_all("td")[0].text:
            # 루프를 돌면서 태그값을 객체에 저장한다.
            comName = table.find_all("tr")[0].find_all("td")[0].text
            comTel = table.find_all("tr")[1].find_all("td")[0].text
            comAddr = table.find_all("tr")[2].find_all("td")[0].text
            comRegiNum = table.find_all("tr")[3].find_all("td")[0].text
            comRegiOrg = table.find_all("tr")[4].find_all("td")[0].text

            # 3.출력
            # 엑셀에 출력
            # 행단위로 추가
            ws.append([i, comName, comTel, comAddr, comRegiNum, comRegiOrg])
            # 100 단위로 세이브
            if i % 100 == 0:
                wb.save("업체목록.xlsx")
            # DB에 출력

        else:
            # print("데이터 없음")
            logger.info("데이터 없음")

    # wb.save("업체목록.xlsx")

    except IndexError as e:
        logger.info(e)
        pass

    finally:
        wb.save("업체목록.xlsx")
