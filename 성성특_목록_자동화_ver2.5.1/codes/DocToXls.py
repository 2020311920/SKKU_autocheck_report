import win32com.client
import sys

def get_doc_text(filename):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(filename)
    text = ""
    for para in doc.Paragraphs:
        text += para.Range.Text + "\n"
    doc.Close()
    word.Quit()
    return text

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python DocToXls.py <doc_file>")
        sys.exit(1)
    
    filename = sys.argv[1]
    text = get_doc_text(filename)

    sample=""
    department = ""
    student_id = ""
    name = ""
    lecture_names = []
    lecture_links = []
    found_department = False
    found_student_id = False
    found_name = False
    found_lecture_name = False
    found_lecture_links = False
    department_bugfix = True

    lines = text.split("\n")
    for i, line in enumerate(lines):
        
        if found_department:
            department = line.strip()[:-1] # 마지막 문자 제거
            found_department = False
            department_bugfix = False
                
        if found_student_id:
            student_id = line.strip()[:-1] # 마지막 문자 제거
            found_student_id = False
        if found_name:
            name = line.strip()[:-1] # 마지막 문자 제거
            found_name = False
        if found_lecture_name:
            lecture_names.append(line.strip())
            #lecture_names.append(lines[i + 3].strip()[:-1])  # 다음 줄도 함께 추가
            found_lecture_name = False

        if found_lecture_links:
            lecture_links.append(line.strip())
            #lecture_links.append(lines[i+3].strip()[:-1])
            found_lecture_links = False

        if "학과" in line:
            if department_bugfix:
                found_department = True
        elif "학번" in line:
            found_student_id = True
        elif "이름" in line:
            found_name = True
        elif "강의명" in line:
            found_lecture_name = True
        elif "링크" in line:
            found_lecture_links = True

    # 결과 출력
    print(department)
    print(student_id)
    print(name+"\n")
    print("강의 코드\n")
    for i in range(0, len(lecture_names), 2):  # 두 줄씩 출력
        print(lecture_links[i])
        print(lecture_links[i+1])
