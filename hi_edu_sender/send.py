"""
send.py — 하이에듀 자동 로그인 + 학부모 문자 발송
사용법: python send.py

send_data.json 파일이 같은 폴더에 있어야 합니다.
(웹 앱의 "send_data.json 다운로드" 버튼으로 생성)
"""

import json
import os
import time
from dotenv import load_dotenv

load_dotenv()

HI_EDU_URL    = os.getenv("HI_EDU_URL", "https://www.hi-edu.or.kr")
HI_EDU_ID     = os.getenv("HI_EDU_ID")
HI_EDU_PW     = os.getenv("HI_EDU_PW")
DATA_FILE     = os.path.join(os.path.dirname(__file__), "send_data.json")

# ── 드라이버 초기화 ──────────────────────────────────
def get_driver(headless=False):
    """
    Chrome WebDriver를 초기화합니다.
    headless=False: 화면이 보이는 모드 (처음 실행 시 권장)
    headless=True:  백그라운드 실행
    """
    try:
        import undetected_chromedriver as uc
        options = uc.ChromeOptions()
        if headless:
            options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--window-size=1280,900")
        driver = uc.Chrome(options=options)
        return driver
    except ImportError:
        # undetected-chromedriver 없으면 일반 selenium 사용
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from webdriver_manager.chrome import ChromeDriverManager
        from selenium.webdriver.chrome.service import Service

        options = Options()
        if headless:
            options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--window-size=1280,900")
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)


# ── 하이에듀 자동화 클래스 ───────────────────────────
class HiEduSender:
    """
    하이에듀(hi-edu.or.kr) 자동 로그인 + 문자 발송.

    ⚠️  주의: 하이에듀 UI 구조는 학교마다 다를 수 있습니다.
        처음 실행할 때는 headless=False 로 화면을 보면서
        셀렉터(CSS/XPath)를 확인하고 아래 코드를 수정하세요.
    """

    def __init__(self, headless=False):
        self.driver = get_driver(headless=headless)
        self.wait   = None
        self._init_wait()

    def _init_wait(self):
        from selenium.webdriver.support.ui import WebDriverWait
        self.wait = WebDriverWait(self.driver, 15)

    # ── 로그인 ──────────────────────────────────────
    def login(self):
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC

        print(f"🌐  하이에듀 로그인 시도... ({HI_EDU_URL})")
        self.driver.get(HI_EDU_URL)
        time.sleep(2)

        # TODO: 실제 셀렉터로 교체 필요
        # 아래는 일반적인 ID/PW 폼 로그인 예시입니다.
        # DevTools(F12) > Elements 탭에서 아이디/비밀번호 input의
        # id 또는 name 속성을 확인해서 교체하세요.
        try:
            id_input = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    'input[type="text"][name*="id"], input#userId, input[name="userId"]'))
            )
            id_input.clear()
            id_input.send_keys(HI_EDU_ID)

            pw_input = self.driver.find_element(By.CSS_SELECTOR,
                'input[type="password"][name*="pw"], input#userPw, input[name="userPw"]')
            pw_input.clear()
            pw_input.send_keys(HI_EDU_PW)

            login_btn = self.driver.find_element(By.CSS_SELECTOR,
                'button[type="submit"], input[type="submit"], .login-btn, #loginBtn')
            login_btn.click()

            time.sleep(2)
            print("✅  로그인 성공")
            return True

        except Exception as e:
            print(f"❌  로그인 실패: {e}")
            print("💡  headless=False 로 실행해서 화면을 확인하세요.")
            return False

    # ── 문자 발송 ────────────────────────────────────
    def send_message(self, phone: str, message: str, student_name: str):
        """
        하이에듀 문자 발송 기능을 자동화합니다.

        TODO: 실제 하이에듀 문자 발송 페이지 URL과 셀렉터를 파악 후 교체.
              1. 브라우저로 하이에듀 접속 → 문자 발송 메뉴 클릭
              2. DevTools > Network 탭에서 API 요청 URL 확인
              3. 또는 Elements 탭에서 수신번호/내용 입력란 셀렉터 확인
        """
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC

        print(f"  📨  {student_name} ({phone}) 발송 중...")

        try:
            # 문자 발송 페이지로 이동 (URL 확인 필요)
            # self.driver.get(f"{HI_EDU_URL}/sms/send")
            # time.sleep(1)

            # 수신번호 입력
            # phone_input = self.wait.until(EC.presence_of_element_located(
            #     (By.CSS_SELECTOR, 'input[name="receiverPhone"], #receiverPhone')
            # ))
            # phone_input.clear()
            # phone_input.send_keys(phone)

            # 메시지 입력
            # msg_input = self.driver.find_element(By.CSS_SELECTOR,
            #     'textarea[name="smsContent"], #smsContent')
            # msg_input.clear()
            # msg_input.send_keys(message)

            # 발송 버튼 클릭
            # send_btn = self.driver.find_element(By.CSS_SELECTOR, '.send-btn, #sendBtn')
            # send_btn.click()
            # time.sleep(1)

            # ─── 임시: 실제 구현 전까지 콘솔에 출력 ───────────
            print(f"     [미구현] 수신: {phone}")
            print(f"     내용: {message[:50]}...")
            # ────────────────────────────────────────────────

            return True

        except Exception as e:
            print(f"     ❌ 발송 실패: {e}")
            return False

    # ── 일괄 발송 ────────────────────────────────────
    def run_batch(self, data: list):
        success, fail = 0, 0
        for item in data:
            phone   = item.get("parent_phone", "")
            message = item.get("message", "")
            name    = item.get("student_name", "")

            if not phone or not message:
                print(f"  ⚠️  {name}: 전화번호 또는 메시지 없음. 건너뜀.")
                continue

            ok = self.send_message(phone, message, name)
            if ok:
                success += 1
            else:
                fail += 1
            time.sleep(0.8)   # 발송 간격

        return success, fail

    def close(self):
        try:
            self.driver.quit()
        except:
            pass


# ── 메인 ──────────────────────────────────────────────
def main():
    if not HI_EDU_ID or not HI_EDU_PW:
        print("❌  .env 파일에 HI_EDU_ID와 HI_EDU_PW를 설정하세요.")
        return

    if not os.path.exists(DATA_FILE):
        print(f"❌  {DATA_FILE} 파일이 없습니다.")
        print("    웹 앱에서 'send_data.json 다운로드' 후 이 폴더에 넣어주세요.")
        return

    with open(DATA_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    print(f"📋  총 {len(data)}명에게 발송 예정")
    for item in data:
        print(f"    • {item['student_no']}번 {item['student_name']} → {item['parent_phone']}")
    print()

    confirm = input("위 목록으로 발송하시겠습니까? (y/N): ").strip().lower()
    if confirm != 'y':
        print("취소되었습니다.")
        return

    sender = HiEduSender(headless=False)  # 처음엔 headless=False로 화면 확인
    try:
        if not sender.login():
            print("로그인 실패로 발송을 중단합니다.")
            return

        success, fail = sender.run_batch(data)
        print()
        print(f"✅  발송 완료: {success}건 성공, {fail}건 실패")

    finally:
        sender.close()


if __name__ == "__main__":
    main()
