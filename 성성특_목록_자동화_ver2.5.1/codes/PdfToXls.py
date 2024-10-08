import PyPDF2
import os
import sys
import re

def extract_text_from_pdf(pdf_path, start_page=None, end_page=None):
    count=1
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""

        if not start_page:
            start_page = 0
        else:
            start_page -= 1  

        if not end_page or end_page > len(reader.pages):
            end_page = len(reader.pages)
        
        for page_num in range(start_page, end_page):
            page_text = reader.pages[page_num].extract_text()
            #print(count, end=" : ")
            #print(page_text)
            #count+=1
            text += page_text
            #this is for test
    
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

    #lines = text.split(" ","\n")
    lines = re.split(r'\s+|\n', text)

    for i, line in enumerate(lines):
        print(i, end=" : ")
        print(line)
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
    #print(department)
    #print(student_id)
    #print(name+"\n")
    #print("강의 코드\n")
    #for i in range(0, len(lecture_names), 2):  # 두 줄씩 출력
        #print(lecture_links[i])
        #print(lecture_links[i+1])

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python PdfToXls.py <Pdf_file>")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    start_page, end_page = None, None
    
    extract_text_from_pdf(pdf_path, start_page, end_page)
