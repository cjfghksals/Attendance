import openpyxl

def get_valid_input():
    while True:
        search_word = input("학번이나 이름을 입력하고 Enter를 눌러주세요. (끝내려면 '00' 입력): ")
        if search_word.lower() == '00':
            return search_word
        if len(search_word) >= 5 or len(search_word.encode('utf-8')) >= 6:
            return search_word
        else:
            print("잘못 입력하셨습니다. 다시 입력해주세요.\n")

file_path = "C:/Users/신동민/Desktop"
file_name = "test.xlsx"

try:
    file_full_path = file_path + "\\" + file_name
    workbook = openpyxl.load_workbook(file_full_path)
    sheet = workbook.active

    while True:
        search_word = get_valid_input()

        if search_word.lower() == '00':
            break

        found = False

        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if search_word in str(cell):
                    found = True
                    break

        if found:
            print(f"출석되었습니다.\n")
        else:
            print(f"미등록자입니다.\n근무자에게 문의해주세요.\n")

except Exception as e:
    print(f"파일을 열 수 없습니다. 오류 메시지: {e}\n")
