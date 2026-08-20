import os
import win32com.client
import mwclient
import urllib3
import glob
import time
import json
import fitz  # PyMuPDF: pip install pymupdf
import sys
import re    # 문장 정렬을 위한 정규표현식 추가
import pythoncom
from pptx import Presentation

#주요기능
#엑셀, 한글, PDF, PPT 파일을 텍스트로 변환해서 위키에 자동으로 업로드하는 기능
#위키 접속정보는 key 파일을 통해 유연하게 계정별로 사용가능
#위키 분류도 key 파일을 통해 변경가능

# SSL 인증서 경고 무시 (사내 위키 접속용)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# 외부 설정 파일(key.txt) 로드 함수
# ==========================================
def load_config():
    if getattr(sys, 'frozen', False):
        current_path = os.path.dirname(os.path.abspath(sys.executable))
    else:
        current_path = os.path.dirname(os.path.abspath(__file__))
    
    config_path = os.path.join(current_path, 'key.txt')
    
    if not os.path.exists(config_path):
        print(f"❌ 설정 파일(key.txt)을 찾을 수 없습니다.")
        print(f"📍 다음 위치에 key.txt를 놓아주세요: {current_path}")
        input("\n종료하려면 엔터를 누르세요...")
        exit()
        
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"❌ key.txt 읽기 오류: {e}")
        input("\n종료하려면 엔터를 누르세요...")
        exit()

# 전역 설정 로드
WIKI_CONFIG = load_config()

def get_base_path():
        if getattr(sys, 'frozen', False):
        # exe 파일로 빌드된 경우
            return os.path.dirname(sys.executable)
        else:
        # 일반 파이썬 스크립트로 실행되는 경우
            return os.path.dirname(os.path.abspath(__file__))
        base_path = get_base_path()

        proc = IntegratedAutomation(current)
class IntegratedAutomation:

    

    def __init__(self, base_path):
        self.folder_path = base_path

        print(f"📂 작업 경로: {self.folder_path}")



    #def __init__(self, folder_path):
        #self.folder_path = os.path.abspath(folder_path)
        #pythoncom.CoInitialize() # COM 객체 안정화
        #print(f"📂 작업 경로: {self.folder_path}\n")

    # 1. 엑셀 -> PDF
    def task_excel(self):
        print("--- [Step 1] 엑셀 변환 시작 ---")
        files = [f for f in os.listdir(self.folder_path) if f.lower().endswith((".xlsx", ".xls"))]
        if not files: return
        excel = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            for f in files:
                try:
                    wb = excel.Workbooks.Open(os.path.join(self.folder_path, f))
                    wb.ExportAsFixedFormat(0, os.path.join(self.folder_path, f"{os.path.splitext(f)[0]}.pdf"))
                    wb.Close(False)
                    print(f"✅ Excel 완료: {f}")
                except Exception as e: print(f"❌ {f} 변환 실패: {e}")
        finally:
            if excel: excel.Quit()

    # 2. 한글 -> PDF
    def task_hwp(self):
        print("\n--- [Step 2] 한글 변환 시작 ---")
        files = [f for f in os.listdir(self.folder_path) if f.lower().endswith((".hwpx", ".hwp"))]
        if not files: return
        hwp = None
        try:       
            hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
            for f in files:
                try:
                    hwp.Open(os.path.join(self.folder_path, f))
                    pdf_p = os.path.join(self.folder_path, f"{os.path.splitext(f)[0]}.pdf")
                    hwp.SaveAs(pdf_p, "PDF")
                    print(f"✅ HWP 완료: {f}")
                except Exception as e: print(f"❌ {f} 변환 실패: {e}")
        finally:
            if hwp: hwp.Quit()

    # 3. PPT -> PDF
    def task_ppt(self):
        print("\n--- [Step 3] PPT 변환 시작 ---")
        files = [f for f in os.listdir(self.folder_path) if f.lower().endswith((".pptx", ".ppt"))]
        if not files: return
        ppt = None
        try:
            # win32com으로 통일하여 안정성 확보
            ppt = win32com.client.DispatchEx("Powerpoint.Application")
            for f in files:
                try:
                    full_path = os.path.join(self.folder_path, f)
                    pdf_path = os.path.join(self.folder_path, f"{os.path.splitext(f)[0]}.pdf")
                    deck = ppt.Presentations.Open(full_path, WithWindow=False)
                    deck.SaveAs(pdf_path, 32) # 32: PDF 포맷
                    deck.Close()
                    print(f"✅ PPT 완료: {f}")
                except Exception as e: print(f"❌ {f} 변환 실패: {e}")
        finally:
            if ppt: ppt.Quit()

    # 4. PDF -> 텍스트 추출 및 문장 정렬 (페이지 번호 삽입 및 출력 보완)
    def task_pdf_to_txt(self):
        print("\n--- [Step 4] PDF 텍스트 추출 및 정렬 시작 ---")
        files = [f for f in os.listdir(self.folder_path) if f.lower().endswith(".pdf")]
        
        for f_name in files:
            try:
                doc = fitz.open(os.path.join(self.folder_path, f_name))
                total_pages = len(doc)
                print(f"📄 대상 파일: {f_name} (총 {total_pages}페이지)")
                
                full_text_list = []

                for i, page in enumerate(doc):
                    page_num = i + 1
                    # 콘솔에 각 페이지 처리 상태 출력
                    print(f"   > [Page {page_num:02d}] 추출 중...")
                    
                    # 텍스트 파일 내부에 저장될 페이지 구분자 (가독성을 위해 상하 개행 추가)
                    full_text_list.append(f"\n\n==================== [ PAGE {page_num} ] ====================\n")
                    
                    # 블록(문단) 단위 추출로 문맥 유지 [cite: 1, 2]
                    blocks = page.get_text("blocks")
                    for b in blocks:
                        block_text = b[4].strip()
                        if block_text:
                            # 줄바꿈 제거 및 공백 정규화 (문장 깨짐 방지)
                            cleaned_block = re.sub(r'\n', ' ', block_text)
                            cleaned_block = re.sub(r'\s+', ' ', cleaned_block).strip()
                            full_text_list.append(cleaned_block)
                
                doc.close()

                if full_text_list:
                    txt_path = os.path.join(self.folder_path, f"{os.path.splitext(f_name)[0]}.txt")
                    with open(txt_path, "w", encoding="utf-8") as tf:
                        for line in full_text_list:
                            # 구분선은 그대로 쓰고, 일반 문장은 마침표 기준으로 개행 처리
                            tf.write(line + "\n")
                            # 문장이 마침표(.)나 기호로 끝나면 가독성을 위해 추가 빈 줄 삽입
                            if not line.startswith("===") and line.endswith(('.', '?', '!')):
                                tf.write("\n")
                    
                    print(f"   ✅ {f_name} 변환 완료 -> {os.path.basename(txt_path)}\n")
                else:
                    print(f"   ⚠️ {f_name}: 추출된 내용이 없습니다.")
            except Exception as e: 
                print(f"   ❌ {f_name} 처리 중 오류 발생: {e}")

    # 5. 미디어위키 일괄 업로드
    def task_wiki_upload(self):
        print("\n--- [Step 5] 미디어위키 업로드 시작 ---")
        all_txt = glob.glob(os.path.join(self.folder_path, "*.txt"))
        txt_files = [f for f in all_txt if os.path.basename(f).lower() != 'key.txt']
        
        if not txt_files:
            print(">> 업로드할 데이터 파일이 없습니다.")
            return

        try:
            site = mwclient.Site(WIKI_CONFIG['SITE_URL'], path=WIKI_CONFIG['PATH'], connection_options={'verify': False})
            site.login(WIKI_CONFIG['USERNAME'], WIKI_CONFIG['PASSWORD'])
            print(f"🔓 로그인 성공: {WIKI_CONFIG['USERNAME']} (@{WIKI_CONFIG['SITE_URL']})")

            category = WIKI_CONFIG.get('CATEGORY_NAME', '제안서')
            print(f"🏷️ 적용될 분류명: {category}")

            for f_path in txt_files:
                title = os.path.splitext(os.path.basename(f_path))[0]
                try:
                    with open(f_path, 'r', encoding='utf-8') as f: 
                        content = f.read()

                    # 본문 + 동적 분류 태그 추가
                    full_body = f"{content}\n\n[[분류:{category}]]"
                    
                    page = site.pages[title]
                    page.save(full_body, summary='통합 자동 업로드 스크립트 실행')
                    print(f"✨ 업로드 완료: {title}")
                    time.sleep(0.3)
                except Exception as file_e:
                    print(f"❌ {title} 업로드 중 오류: {file_e}")
                    
        except Exception as e:
            print(f"❌ 위키 접속/로그인 오류: {e}")

if __name__ == "__main__":
# 수정 전: current = os.path.dirname(os.path.abspath(__file__))
    
    # 수정 후: 빌드 여부에 따라 실제 경로를 잡도록 변경
    if getattr(sys, 'frozen', False):
        # .exe로 실행 중일 때 (.exe 파일이 있는 폴더)
        current = os.path.dirname(os.path.abspath(sys.executable))
    else:
        # .py로 실행 중일 때 (스크립트 파일이 있는 폴더)
        current = os.path.dirname(os.path.abspath(__file__))
    
    # 확인용 출력 (빌드 후 실행했을 때 Temp가 아닌 실제 폴더가 나와야 함)
    print(f"📍 현재 인식된 작업 경로: {current}")
    
    proc = IntegratedAutomation(current)
    
    # 전체 프로세스 순차 실행
    proc.task_excel()
    proc.task_hwp()
    proc.task_ppt()
    proc.task_pdf_to_txt()
    proc.task_wiki_upload()
    
    print("\n🎉 모든 작업이 완료되었습니다.")
    input("종료하려면 엔터를 누르세요...")