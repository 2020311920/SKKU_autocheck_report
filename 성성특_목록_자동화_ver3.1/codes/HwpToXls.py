# test.py 스크립트

import sys
import olefile
import zlib
import struct

def get_hwp_text(filename):
    try:
        f = olefile.OleFileIO(filename)
        dirs = f.listdir()

        # HWP 파일 검증
        if ["FileHeader"] not in dirs or \
           ["\x05HwpSummaryInformation"] not in dirs:
            raise Exception("Not Valid HWP.")

        # 문서 포맷 압축 여부 확인
        header = f.openstream("FileHeader")
        header_data = header.read()
        is_compressed = (header_data[36] & 1) == 1

        # Body Sections 불러오기
        nums = []
        for d in dirs:
            if d[0] == "BodyText":
                nums.append(int(d[1][len("Section"):]))
        sections = ["BodyText/Section"+str(x) for x in sorted(nums)]

        # 전체 text 추출
        text = ""
        for section in sections:
            bodytext = f.openstream(section)
            data = bodytext.read()
            if is_compressed:
                unpacked_data = zlib.decompress(data, -15)
            else:
                unpacked_data = data
        
            # 각 Section 내 text 추출    
            section_text = ""
            i = 0
            size = len(unpacked_data)
            while i < size:
                header = struct.unpack_from("<I", unpacked_data, i)[0]
                rec_type = header & 0x3ff
                rec_len = (header >> 20) & 0xfff

                if rec_type in [67]:
                    try:
                        rec_data = unpacked_data[i+4:i+4+rec_len]
                        section_text += rec_data.decode('utf-16')
                        section_text += "\n"
                    except UnicodeDecodeError:
                        # cp949 오류 처리
                        section_text += rec_data.decode('cp949', errors='replace')
                        section_text += "\n"

                i += 4 + rec_len

            text += section_text
            text += "\n"

        return text

    except Exception as e:
        raise Exception("Error while processing HWP file: {}".format(str(e)))


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 test.py <hwp_file>")
        sys.exit(1)

    hwp_file_path = sys.argv[1]
    try:
        hwp_text = get_hwp_text(hwp_file_path)
        #print(hwp_text)

        # 학과, 학번, 이름, 강의명이 나오는 위치를 찾고 텍스트 추출
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

        lines = hwp_text.split("\n")
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
        #print("강의명\n")
        print("강의 코드\n")
        for i in range(0, len(lecture_names), 2):  # 두 줄씩 출력
            # print(lecture_names[i])
            print(lecture_links[i])
            # print(lecture_names[i + 1])
            print(lecture_links[i+1])

    except Exception as e:
        print("Error:", e)

