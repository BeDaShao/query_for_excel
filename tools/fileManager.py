import os


def check_not_open_file(query_file_name,files):
    for file in files:
        if file == query_file_name:
            print('請關閉檔案:',query_file_name)
            os.system("pause") 
            return False
    return True
    
def get_file_name(current_gwd):
    # change directory to parent dir of this python file
    DIR_PATH = os.path.join(current_gwd, os.path.pardir)
    os.chdir(DIR_PATH)

    query_file_name = '篩選後資料.xlsx'
    files = os.listdir()
    # check the query file isn't opened
    while True:
        files = os.listdir()
        if check_not_open_file('~$'+ query_file_name, files):
            break
    
    # find the excel file  
    for file in files:
        if '.xlsx' in file and file != query_file_name:
            return file
    return None