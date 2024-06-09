
import win32com.client
import pyautogui
import pandas as pd
import csv


def excel_to_csv(excel_file, csv_file):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(excel_file)
    ws = wb.Sheets(1)
    
    data = ws.UsedRange.Value
    df = pd.DataFrame(data[1:], columns=data[0])
    # df.columns =  ["Year", "Program", "Mailing", "CAP Number", "Kit#", "Kit ID", "Kit Mailed", "Original Evaluation", "Revision Evaluation", "Expedited Evaluation", "Table Type", "Page No.", "Analyte", "Unit of Measure", "Evaluation Result", "Peer Instrument", "Your Instrument", "Peer Method", "Your Method", "Peer Reagent", "Instrument", "Method", "Methods", "Reagent", '''Allowable Error (whichever is greater)''', "Peer Results Summary Table", "Your Peer Group", "Peer Group Size", "Evaluation Type", "Goal for Total Error (TE)", "PeerGroup", "Test", "Specimen", "Your Result", "Mean", "S.D.", "No. of Labs", "S.D.I", "bais %", "bais conc", "LOA Lower", "LOA Upper", "Good Response", "Acceptable Response", "Intended Response", "Your Grade", "Assay 1", "Assay 2", "Your Mean", "Peer Mean", "Peer N", "Assigned Target", "Difference", "Allowable Error", "Range", "%Verified", "%Different", "%Linear", "%Nonlinear", "%Imprecise", "Best-fit Target", "Relative Concentration"]
    df_keys = list(df.keys())
    print('df_keys :', df_keys)
    """ CRLF 처리 """
    for i in range(len(df_keys)):
        df[df_keys[i]] = df[df_keys[i]].str.replace('\n',r'\n')
        df[df_keys[i]] = df[df_keys[i]].str.replace('\r',r'\r')
    """ 콤마(,)가 들어간 데이터가 있어서 구분자를 콤마(,)로 사용할 수 없음 """
    ascii_code = 30 #30 => RECORD SEPARATOR (RS) UP ARROW / #124 => |(vertical bar)
    delimiter = chr(ascii_code)
    df.to_csv(csv_file, index=False, sep=delimiter) 
    """ 데이터마다 쌍따옴표 처리할 때 아래 코드 사용 """
    # df.to_csv(csv_file, sep=delimiter, quotechar='"', quoting=csv.QUOTE_ALL, index=False)

    wb.Close()
    excel.Quit()

daily_csv_file=r"D:\RPA외부정도관리\처리결과\CAP\CSV\DB_CAP(2024-05-09).csv"
excel_file=r"D:\RPA외부정도관리\처리결과\CAP\DB결과파일\DB_CAP(2024-05-09).xlsx"
excel_to_csv(excel_file, daily_csv_file)