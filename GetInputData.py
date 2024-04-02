import win32com.client as win32
#-------------------------------------------------------------------------------
# 폴더내 특정 파일 리스트 추출
#-------------------------------------------------------------------------------
def GetFileList(folder, ext):
    import os
    files = [f.name for f in os.scandir(folder) if f.is_file() and f.name.endswith(ext)]
    print('Num of xlsx files:', len(files))
    return files
#-------------------------------------------------------------------------------
# 지난달 파일명 추출
#-------------------------------------------------------------------------------
def Get전월파일(ExcelList):
    for i in range(len(ExcelList)):
        if '구매팀' in ExcelList[i]:
            전월file = ExcelList[i]
    return 전월file
#-------------------------------------------------------------------------------
# 파일 이동
#-------------------------------------------------------------------------------
import shutil
def MoveFile(src_path, dst_path):
    try:
        shutil.move(src_path, dst_path)
        print("File moved successfully!")
    except IOError as e:
        print(f"Unable to move file. {e}")
#-------------------------------------------------------------------------------
# 컬럼: 문자 -> 숫자
#-------------------------------------------------------------------------------
def ColToNum(colStr):
    """ Convert base26 column string to number. """
    expn = 0
    colNum = 0
    for char in reversed(colStr):
        colNum += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    #return colNum-1
    return colNum
#-------------------------------------------------------------------------------
# 컬럼: 숫자 -> 문자
#-------------------------------------------------------------------------------
def NumToCol(n) :
    string = ""
    #n = n+1
    while n > 0 : 
        n, remainder = divmod(n, 26)
        if(remainder == 0) :
            n = n-1
            string = 'Z' + string   
        else :
            string = chr(ord('A') + remainder -1) + string
    return string
#-------------------------------------------------------------------------------
# 계약서 이름 가져오기
#-------------------------------------------------------------------------------
def Get계약명(ws, row, col):
    return ws.Cells(row, col).text
#-------------------------------------------------------------------------------
# 증빙파일
#-------------------------------------------------------------------------------
def Get증빙파일(ws, row, col):
    증빙파일 = ws.Cells(row, col).text
    return 증빙파일
#-------------------------------------------------------------------------------
# 승입합계 
#-------------------------------------------------------------------------------
def Get승인합계(ws, row, col):
    승인합계s = ws.Cells(row, col).text.replace(',','')
    return 승인합계s
#-------------------------------------------------------------------------------
# 제목키워드
#-------------------------------------------------------------------------------
def Get제목키워드(ws, row, col):
    제목키워드 = ws.Cells(row, col).text
    return 제목키워드
#-------------------------------------------------------------------------------
# 문서번호
#-------------------------------------------------------------------------------
def Get문서번호(ws, row, col):
    문서번호 = ws.Cells(row, col).text
    return 문서번호
#-------------------------------------------------------------------------------
# 안내메일 또는 독촉메일 발행 월
#-------------------------------------------------------------------------------
def Get메일발송월s(ws, row, col):
    발송월 = ws.Cells(row, col).text
#    print('발송월:', 발송월)
    if 발송월 == '매월':
        발송월s = list(range(1,13))
    else:
        t = 발송월.split(',')
        발송월s = [int(x) for x in t]
    return 발송월s
#-------------------------------------------------------------------------------
# 처리결과
#-------------------------------------------------------------------------------
def Get처리결과(ws, row, col):
	dic = {}
	for idx in range(3):
		key = ws.Cells(row+idx, col).text
		dic[key] = ws.Cells(row+idx, col+1).text
	return dic
#-------------------------------------------------------------------------------
# 이번 달 추출
#-------------------------------------------------------------------------------
def get_current_month():
    import datetime
    current_month = datetime.date.today().month
    return current_month
#-------------------------------------------------------------------------------
# 이번 달의 1일 및 말일 추출
#-------------------------------------------------------------------------------
def get_first_last_day():
    from datetime import datetime, timedelta
    # 오늘날짜
    today = datetime.now().date()
    # 지난 달 1일 
    cur_date = datetime(today.year, today.month, 1) - timedelta(days=1)
    first_day = cur_date.strftime("%Y%m01")
    # 다음달 지정
    if today.month == 12:
        next_month = 1
        next_year = today.year + 1
    else:
        next_month = today.month + 1
        next_year = today.year
    # 달의 마지막 day
    last_day = (datetime(next_year, next_month, 1) - timedelta(days=1)).strftime("%Y%m%d")
    return first_day, last_day
#-------------------------------------------------------------------------------
# 매입전자세금계산서 기표처리(간략/상세)
#-------------------------------------------------------------------------------
def Get기표처리(ws, rows, col):
    dic = {}
    for row in rows:
        if row == 10: #전자계산서일자(시작)
            key = ws.Cells(row, col).text
            first_day, last_day = get_first_last_day()
            dic[key] = str(int(first_day))
        elif row == 11: #전자계산서일자(종료)
            key = ws.Cells(row, col).text
            first_day, last_day = get_first_last_day()
            dic[key] = str(int(last_day))
        else:
            key = ws.Cells(row, col).text
            dic[key] = ws.Cells(row, col+1).text
    return dic
#-------------------------------------------------------------------------------
# 코스트센터별 금액 분배
#-------------------------------------------------------------------------------
def Get코센별분배(ws, row, cols):
    keys = [ws.Cells(row, col).text for col in cols]
    v_dic= {}
    for col, key in zip(cols, keys):
        lines  = []
        inner_row = row+1
        while True:
            cell = ws.Cells(inner_row, col).text
            if cell == '':
                break
            else:
                lines.append(cell.replace(',',''))
            inner_row += 1
        v_dic[key] = lines
    return v_dic
#-------------------------------------------------------------------------------
# 입력데이터 자료구조로 읽어오기
#-------------------------------------------------------------------------------
def Get입력데이터(ws, total_case):
    input_data = []
    for idx in list(range(total_case)):
        line = {}
        #계약서명 
        row = 1
        col = ColToNum('B') + idx*4
        line['계약명'] = Get계약명(ws, row, col)
        #증빙파일 
        row = 2
        col = ColToNum('C') + idx*4
        line['증빙파일'] = Get증빙파일(ws, row, col)
        #승인합계
        row = 3
        col = ColToNum('C') + idx*4
        line['승인합계'] = Get승인합계(ws, row, col)
        #제목키워드
        row = 4
        col = ColToNum('C') + idx*4
        line['제목키워드'] = Get제목키워드(ws, row, col)
        #문서번호
        row = 5
        col = ColToNum('C') + idx*4
        line['문서번호'] = Get문서번호(ws, row, col)
        #메일발송월
        row = 6
        col = ColToNum('C') + idx*4
        line['메일발송월'] = Get메일발송월s(ws, row, col)
        # print("line['메일발송월']:", line['메일발송월'])
        #처리결과
        row = 8
        col = ColToNum('D') + idx*4
        dic = Get처리결과(ws, row, col)
        line['처리결과'] = dic
        #매입전자세금계산서 기표처리(간략)
        rows = list(range(8,14))
        col = ColToNum('B') + idx*4
        line['기표처리간략'] = Get기표처리(ws, rows, col) 
        #매입전자세금계산서 기표처리(상세)
        rows = list(range(15,27))
        col = ColToNum('B') + idx*4
        line['기표처리상세'] = Get기표처리(ws, rows, col) 
        #코스트센터별 금액 분배
        row = 27
        start_col = ColToNum('B') + idx*4
        end_col = ColToNum('E') + idx*4
        cols = list(range(start_col, end_col+1))
        line['코센별분배'] = Get코센별분배(ws, row, cols)
        input_data.append(line)
    return input_data
