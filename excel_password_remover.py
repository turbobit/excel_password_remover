# 엑셀 파일 이름을 을 변수로 받아 암호 걸릴걸 제거하고 저장

import msoffcrypto
import io
import os
import sys
import itertools
import string
import time
from datetime import datetime
import math
import signal

# 진행상황 표시를 위한 상수
SPINNER_CHARS = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"
PROGRESS_BAR_LENGTH = 30

# 전역 변수로 종료 플래그 추가
should_exit = False

def signal_handler(signum, frame):
    """
    시그널 핸들러: Ctrl+C 등의 시그널을 처리합니다.
    """
    global should_exit
    should_exit = True
    print("\n\n프로그램을 종료합니다...")
    sys.exit(0)

def get_spinner_char(index):
    """
    스피너 문자를 반환합니다.
    """
    return SPINNER_CHARS[index % len(SPINNER_CHARS)]

def create_progress_bar(progress, length=PROGRESS_BAR_LENGTH):
    """
    진행률 바를 생성합니다.
    """
    filled_length = int(length * progress)
    bar = '█' * filled_length + '░' * (length - filled_length)
    return f"[{bar}] {progress:.1%}"

def try_password(office_file, password):
    """
    주어진 비밀번호로 파일을 해제해봅니다.
    """
    try:
        temp_buffer = io.BytesIO()
        office_file.load_key(password=password)
        office_file.decrypt(temp_buffer)
        return True
    except:
        return False

def generate_passwords(min_length=4, max_length=6):
    """
    가능한 모든 비밀번호 조합을 생성합니다.
    """
    chars = string.ascii_letters + string.digits
    for length in range(min_length, max_length + 1):
        for password in itertools.product(chars, repeat=length):
            yield ''.join(password)

def format_time(seconds):
    """
    초 단위 시간을 읽기 쉬운 형식으로 변환합니다.
    """
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f"{int(hours)}시간 {int(minutes)}분 {int(seconds)}초"

def find_password(input_file, min_length=4, max_length=6):
    """
    파일의 비밀번호를 찾습니다.
    """
    global should_exit
    try:
        print(f"\n=== 비밀번호 찾기 시작 ===")
        print(f"시작 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"비밀번호 길이: {min_length}자리 ~ {max_length}자리")
        print("========================\n")

        # 전체 조합 수 계산
        chars = string.ascii_letters + string.digits
        total_combinations = sum(len(chars) ** i for i in range(min_length, max_length + 1))

        with open(input_file, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            
            if not office_file.is_encrypted():
                print(f"알림: 파일이 이미 암호화되어 있지 않습니다: {input_file}")
                return None

            # 진행상황 모니터링
            start_time = time.time()
            last_update_time = time.time()
            attempts = 0
            spinner_index = 0
            current_password = ''

            try:
                for password in generate_passwords(min_length, max_length):
                    if should_exit:
                        print("\n\n사용자에 의해 중단되었습니다.")
                        return None

                    attempts += 1
                    current_password = password
                    progress = attempts / total_combinations

                    if try_password(office_file, password):
                        end_time = time.time()
                        total_time = end_time - start_time
                        
                        print(f"\n\n=== 비밀번호 찾기 성공! ===")
                        print(f"비밀번호: {password}")
                        print(f"시도 횟수: {attempts:,}회")
                        print(f"소요 시간: {format_time(total_time)}")
                        print(f"종료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                        print("========================\n")
                        return password

                    # 0.05초마다 진행상황 출력
                    current_time = time.time()
                    if current_time - last_update_time >= 0.05:
                        elapsed_time = current_time - start_time
                        spinner = get_spinner_char(spinner_index)
                        spinner_index += 1

                        # 화면 지우기
                        os.system('cls' if os.name == 'nt' else 'clear')
                        
                        # 진행률 바 표시
                        print(f"\n{spinner} 진행률: {create_progress_bar(progress)}")
                        print(f"시도 횟수: {attempts:,}회 | 경과 시간: {format_time(elapsed_time)}")
                        print(f"\n현재 시도 중인 비밀번호: {current_password}")
                        
                        last_update_time = current_time

            except KeyboardInterrupt:
                print("\n\n사용자에 의해 중단되었습니다.")
                return None

            print("\n\n=== 비밀번호 찾기 실패 ===")
            print(f"시도 횟수: {attempts:,}회")
            print(f"소요 시간: {format_time(time.time() - start_time)}")
            print(f"종료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print("========================\n")
            return None

    except Exception as e:
        print(f"오류 발생: {str(e)}")
        return None

def remove_excel_password(input_file, output_file=None, password=None):
    """
    엑셀 파일의 암호를 제거하고 새로운 파일로 저장합니다.
    """
    try:
        if not os.path.exists(input_file):
            print(f"오류: 파일을 찾을 수 없습니다: {input_file}")
            return False

        if output_file is None:
            file_name = os.path.basename(input_file)
            output_file = os.path.join(os.path.dirname(input_file), f"unlocked_{file_name}")

        with open(input_file, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            
            if not office_file.is_encrypted():
                print(f"알림: 파일이 이미 암호화되어 있지 않습니다: {input_file}")
                return False

            temp_buffer = io.BytesIO()
            if password:
                office_file.load_key(password=password)
            else:
                office_file.load_key(password="")
            office_file.decrypt(temp_buffer)

            with open(output_file, 'wb') as f:
                f.write(temp_buffer.getvalue())

        print(f"성공: 암호가 제거된 파일이 저장되었습니다: {output_file}")
        return True

    except Exception as e:
        print(f"오류 발생: {str(e)}")
        return False

def main():
    # 시그널 핸들러 등록
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    if len(sys.argv) < 2:
        print("사용법: python index.py <엑셀파일경로> [출력파일경로]")
        return

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    # 항상 비밀번호 찾기 모드로 실행
    password = find_password(input_file)
    if password:
        remove_excel_password(input_file, output_file, password)
    else:
        print("비밀번호를 찾지 못했습니다. 프로그램을 종료합니다.")

if __name__ == "__main__":
    main()