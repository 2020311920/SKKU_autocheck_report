try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
import sys

NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = NAMESPACE + 'p'
TEXT = NAMESPACE + 't'


def get_docx_text(filename):
    document = zipfile.ZipFile(filename)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)
    paragraphs = []
    for paragraph in tree.iter(PARA):
        texts = [node.text
                for node in paragraph.iter(TEXT)
                if node.text]
        if texts:
            paragraphs.append(''.join(texts))
    return '\n\n'.join(paragraphs)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python DocxToXls.py <docx_file>")
        sys.exit(1)
    
    filename = sys.argv[1]
    text = get_docx_text(filename)
    
    keywords = ["학과", "학번", "이름", "링크"]
    
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

    sample_depart_found=False

    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i%2!=0:
            continue
        
        
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
            #lecture_names.append(lines[i + 2].strip())  # 다음 줄도 함께 추가
            found_lecture_name = False

        if found_lecture_links:
            lecture_links.append(line.strip())
            #lecture_links.append(lines[i+2].strip())
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
  

