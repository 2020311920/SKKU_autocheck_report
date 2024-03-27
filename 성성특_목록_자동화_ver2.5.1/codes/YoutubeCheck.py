import sys
import re
import warnings

def extract_youtube_id(url):
    # 유효한 유튜브 링크 패턴 정의
    youtube_regex = (
        r'(https?://)?(www\.)?'
        '(youtube|youtu|youtube-nocookie)\.(com|be)/'
        '(watch\?v=|embed/|v/|.+\?v=)?([^&=%\?]{11})')

    # 정규식을 이용하여 링크에서 고유 코드 추출
    youtube_match = re.match(youtube_regex, url)
    if youtube_match:
        return youtube_match.group(6)
    else:
        return -1

def is_valid_youtube_url(url):
    # 유효한 유튜브 링크 패턴 정의
    youtube_regex = (
        r'(https?://)?(www\.)?'
        '(youtube|youtu|youtube-nocookie)\.(com|be)/'
        '(watch\?v=|embed/|v/|.+\?v=)?([^&=%\?]{11})')

    # 정규식을 이용하여 링크가 유효한지 확인
    return re.match(youtube_regex, url) is not None

if __name__ == "__main__":
    # 경고 무시
    warnings.filterwarnings("ignore")

    # 커맨드 라인에서 유튜브 링크를 입력 받음
    if len(sys.argv) != 2:
        print("Usage: python code.py <youtube_link>")
        sys.exit(1)

    youtube_url = sys.argv[1]

    if is_valid_youtube_url(youtube_url):
        video_id = extract_youtube_id(youtube_url)
        if video_id != -1:
            #print("유효한 유튜브 링크입니다.")
            print(video_id, end="")
        else:
            print("유효한 유튜브 링크지만 고유 코드를 추출할 수 없습니다.", end="")
    else:
        print("유효하지 않은 유튜브 링크입니다.", end="")
