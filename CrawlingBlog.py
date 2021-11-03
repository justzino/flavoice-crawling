import os
import time

import chromedriver_autoinstaller
import xlsxwriter
from dotenv import load_dotenv
from selenium import webdriver

env_path = '.env'
load_dotenv(dotenv_path=env_path, verbose=True)
env = os.getenv
mode = 0

# webdirver 설정(Chrome, Firefox 등)
chromedriver_autoinstaller.install()

driver = webdriver.Chrome()
driver.set_window_size(800, 800)

version = 0
sleepTime = 0.3
resetServerCounter = 5  # 몇번 접근하여 서버 시간 불러올지 횟수 정함
serverCounter = resetServerCounter
currentTime = None


def save_to_xls(workbook):
    kor_to_note = {'도': 'C', '도#': 'C#', '레': 'D', '레#': 'D#', '미': 'E', '파': 'F',
                   '파#': 'F#', '솔': 'G', '솔#': 'G#', '라': 'A', '라#': 'A#', '시': 'B'}

    worksheet = workbook.add_worksheet()

    worksheet.set_column('A:A', 30)  # A 열의 너비를 40으로 설정
    worksheet.set_column('B:B', 12)  # B 열의 너비를 12로 설정

    worksheet.write(0, 0, '노래 제목')
    worksheet.write(0, 1, 'max_pitch')
    worksheet.write(0, 2, '노래 설명')
    worksheet.write(0, 3, '가수')
    worksheet.write(0, 4, '데뷔일')
    worksheet.write(0, 5, '장르')

    excel_row = 2

    for i in range(5, 606):
        line = driver.find_element_by_xpath(f'/html/body/div[6]/div[1]/div[4]/div[2]/p[{i}]').text
        data = list(line.split('  '))
        data_list = []
        for c in data:
            if c == '':
                continue
            while c.startswith(' '):
                c = c[1:]
            data_list.append(c)

        try:
            singer, title, note = data_list[0], data_list[1], data_list[2]
            octave, pitch = str(int(note[0]) + 3), note[3:]
            if len(pitch) > 1:
                if pitch[1] == '#':
                    pitch = pitch[:2]
                else:
                    pitch = pitch[0]
            note = kor_to_note[pitch] + octave

            # 엑셀 저장(텍스트)
            worksheet.write(excel_row, 0, title)
            worksheet.write(excel_row, 1, note)
            worksheet.write(excel_row, 3, singer)

            print(f"singer: {singer},\t title: {title},\t note: {note}")

            # 엑셀 행 증가
            excel_row += 1
        except:
            print("실패: ", data_list)
            continue
    time.sleep(1)


# main
print('페이지 키는 중..')
driver.get('')

# Excel 처리 선언
savePath = env('save_path')
workbook = xlsxwriter.Workbook(savePath + 'result.xlsx')

save_to_xls(workbook)

# 엑셀 파일 닫기
workbook.close()  # 저장

print("저장 완료")
