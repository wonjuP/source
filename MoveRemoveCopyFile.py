import os
import shutil
#----------------------------------------------------------------------------
# folder가 비우기
#----------------------------------------------------------------------------
def RemoveAllFiles(folder):
  for file in os.listdir(folder):
    try:
      os.remove(os.path.join(folder, file))
    except:
      "continue"
#-------------------------------------------------------------------------------
# 
#-------------------------------------------------------------------------------         
def replace_file(folder_name, file_path):
    print('파일 이름 :', file_path.split('\\')[-1])
    if not file_path.split('\\')[-1] in os.listdir(folder_name):
        print("파일 없음")
        shutil.move(file_path, folder_name)
    else:
        print("파일 있음")
        os.remove(folder_name+'\\'+file_path.split('\\')[-1])
        shutil.move(file_path, folder_name)
#-------------------------------------------------------------------------------
# 
#-------------------------------------------------------------------------------      
def create_folder_and_move_files(folder_name, file_path):
    if not os.path.exists(folder_name): #폴더가 없으면
        print("Year 폴더 존재")
        os.makedirs(folder_name)
        replace_file(folder_name, file_path)
    else: #폴더 있으면
        print("Year 폴더 부재")
        replace_file(folder_name, file_path)
        

# from datetime import datetime, timedelta
# import datetime
# date = datetime.date.today().strftime("%Y-%m-%d")
# print('date :', date)

#-------------------------------------------------------------------------------
# daily 액셀 파일 삭제
#------------------------------------------------------------------------------- 
def remove_daily_files(DB폴더):
    collection = os.listdir(DB폴더)
    print(collection)
    for file in collection:
        if '-' in file:
            os.remove(DB폴더+file)
            print(f'{file} 삭제 완료')
