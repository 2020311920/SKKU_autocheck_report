#pip install pdfplumber 필요
import pdfplumber
import sys

def extract_tables_from_pdf(pdf_path):
    text=""
    # PDF 파일 열기
    with pdfplumber.open(pdf_path) as pdf:
        # 각 페이지를 순회하며 표 추출
        for page in pdf.pages:
            # 페이지 내의 모든 표를 추출
            tables = page.extract_tables()
            for table in tables:
                # 각 행을 순회하며 셀 출력
                for row in table:
                    for element in row:
                        if(element!=None):
                            for word in element:
                                if(word!='\n'):
                                    #print(word, end="")
                                    text+=word
                            #print()
                            text+="\n"
    
    return text

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_tables.py <Pdf_file>")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    text=extract_tables_from_pdf(pdf_path)

    #print(text)

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
        #print(i, end=" : ")
        #print(line)
        if found_department:
            department = line.strip()
            found_department = False
            department_bugfix = False
                
        if found_student_id:
            student_id = line.strip()
            found_student_id = False
        if found_name:
            name = line.strip()
            found_name = False
        if found_lecture_name:
            lecture_names.append(line.strip())
            #lecture_names.append(lines[i + 1].strip())  # 다음 줄도 함께 추가
            found_lecture_name = False

        if found_lecture_links:
            lecture_links.append(line.strip())
            #lecture_links.append(lines[i+1].strip())
            
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
