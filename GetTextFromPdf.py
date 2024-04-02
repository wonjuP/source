import fitz
import os
import win32com.client as win
#-------------------------------------------------------------------------------
# 영역(Rectangle Area) 지정하여 텍스트를 읽는 함수 
#-------------------------------------------------------------------------------
def GetTextFromRectangle(x1y1, x2y2, page):
  text = []
  rect = fitz.Rect(x1y1, x2y2)  # define your rectangle here
  text = page.get_textbox(rect)  # get text from rectangle
  return text.strip()   

def GetPDFText(pdfFile, pageNo, rect_coords):
  with fitz.open(pdfFile) as doc: 
    text = GetTextFromRectangle(rect_coords[0], rect_coords[1], doc[pageNo])
    return text
#-------------------------------------------------------------------------------
# Input: pdf파일의 경로
# Output: WordsDic - word 단위의 텍스트와 해당 좌표의 모음
#         BlocksDic - block 단위의 텍스트와 해당 좌표의 모음 
#-------------------------------------------------------------------------------
def GetWordsBlocks(pathToPDFfile, WordsDic, BlocksDic):
  with fitz.open(pathToPDFfile) as document: 
    for page_number, page in enumerate(document, start=0):
      words = page.get_text("words")
      r_words = []
      for word in words:
        r_words.append((round(word[0],1), round(word[1],1), round(word[2],1), round(word[3],1), word[4], word[5], word[6]))
      WordsDic[page_number] = r_words

    for page_number, page in enumerate(document, start=0):
      words = page.get_text("blocks")
      r_words = []
      for word in words:
        r_words.append((round(word[0],1), round(word[1],1), round(word[2],1), round(word[3],1), word[4], word[5], word[6]))
      BlocksDic[page_number] = r_words
#-------------------------------------------------------------------------------
# 특정 확장자 파일 리스트 추출
#-------------------------------------------------------------------------------
def GetFileList(folder, ext):
  files = [f.name for f in os.scandir(folder) if f.is_file() and f.name.endswith(ext)]
  return files

################################################################################
# 실행코드
################################################################################
config_file = r"C:\RPA\PDF데이터추출\Config.xlsx"
excel = win.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(config_file)

# 유형정보 시트 데이터 추출 -------------
ws = wb.Sheets("유형정보")
used_range = ws.UsedRange
column_a_values = used_range.Columns(1).Value
column_b_values = used_range.Columns(2).Value
type_sep = {}
for i in range(1, len(column_a_values)):
  key = column_a_values[i][0]  # Value from column A
  value = column_b_values[i][0]  # Value from column B
  type_sep[key] = value

# 영역정보 시트 데이터 추출 -------------
ws = wb.Sheets("영역정보")
cnt = ws.UsedRange.Rows.Count
# print('cnt :', cnt) #34
range_info = list(ws.Range("A2:F"+str(cnt)).Value)
for i in range(len(range_info)):
    if range_info[i][0] is None and i > 0:
        range_info[i] = (range_info[i-1][0], range_info[i][1]) + range_info[i][2:]
dic = {}
for item in range_info:
    company_name = item[0].strip()
    if company_name not in dic:
        dic[company_name] = {}
    key = item[1].strip()
    values = item[2:]
    if None in values:
        values = [0.0 for _ in values]  # Replace None with 0.0
    dic[company_name][key] = values
# print('영역정보 dic :', dic['홈텍스']['승인번호'])
    
# 액셀 종료 -------------
wb.Close()
excel.Quit()

#텍스트 추출 -------------
folder = r"C:\RPA\PDF데이터추출"+"\\"  #pdf파일 있는 폴더의 경로
files = GetFileList(folder, '.pdf')
# print('files :', files)
for idx, file in enumerate(files):
  pdf_file = folder+file
  WordsDic = {} 
  BlocksDic = {} 
  GetWordsBlocks(pdf_file, WordsDic, BlocksDic) #Dict형태로 Words/Blocks 텍스트 및 좌표값 추출
  
  # WordsDic/BlocksDic txt파일 쓰기
  text_folder = r"C:\RPA\PDF데이터추출\텍스트파일"+"\\"
  with open(text_folder+file.replace(".pdf","(WordsDic).txt"), "w") as txt:
    txt.write(str(WordsDic))
  with open(text_folder+file.replace(".pdf","(BlocksDic).txt"), "w") as txt:
    txt.write(str(BlocksDic))
  # print('BlocksDic_'+str(idx), BlocksDic)

  for type_key in type_sep.keys():
    seperator = type_sep[type_key]
    # print('seperator :', seperator)
    for block in BlocksDic[0]:
      if seperator in block[4]:
        print("================= ["+ type_key +"] =================")
        for r_name in dic[type_key].keys():
          rec_pos = dic[type_key][r_name]
          # print('영역 좌표정보 :', rec_pos)
          x1=rec_pos[0]; y1=rec_pos[1]; x2=rec_pos[2]+x1; y2 =rec_pos[3]+y1
          x1y1 = (x1,y1); x2y2 = (x2,y2)
          # print('x1y1 :', x1y1, '-', 'x2y2 :', x2y2)
          영역data = GetPDFText(pdf_file, 0, (x1y1,x2y2)).replace('\n', '')
          print(r_name,':', 영역data)
