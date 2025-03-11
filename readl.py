from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import shutil
from datetime import datetime, timedelta
import pandas as pd
import os

chrome_options = Options()
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--headless')
driver = webdriver.Chrome(service=Service(f"./chromedriver-linux64/chromedriver"), options=chrome_options)
driver.maximize_window()
wait = WebDriverWait(driver, 10)

# 엑셀 파일명에 포함할 변수 초기화
period = ""
category_1 = ""
category_2 = ""
category_3 = ""
device_selection = ""
gender_selection = ""
age_selection = ""

data = {
    "인기검색어 순위": [],
    "인기검색어": []
}


def select_period(start_year, start_month, start_day, end_year, end_month, end_day):
    global period
    retry_count = 0  # 최대 재시도 횟수

    while retry_count < 3:
        try:
            print("🔍 종료일 먼저 선택 시작...")
            found_end_year = found_end_month = found_end_day = False
            found_start_year = found_start_month = found_start_day = False

            # ✅ 종료 연도 선택
            end_year_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "(//div[@class='select w2'])[2]//span[@class='select_btn']")))
            driver.execute_script("arguments[0].click();", end_year_btn)
            time.sleep(1)

            options = driver.find_elements(By.XPATH, "//div[@class='select w2']//ul[@class='select_list scroll_cst']//li/a")
            for option in options:
                if option.text.strip() == end_year:
                    driver.execute_script("arguments[0].click();", option)
                    print(f"✅ 종료 연도 {end_year} 선택 완료.")
                    found_end_year = True
                    break

            if not found_end_year:
                print(f"❌ 종료 연도 {end_year} 찾을 수 없음!")

            # ✅ 종료 월 선택
            end_month_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "(//div[@class='select w3'])[3]//span[@class='select_btn']")))
            driver.execute_script("arguments[0].click();", end_month_btn)
            time.sleep(1)

            options = driver.find_elements(By.XPATH, "//div[@class='select w3']//ul[@class='select_list scroll_cst']//li/a")
            for option in options:
                if option.text.strip() == end_month:
                    driver.execute_script("arguments[0].click();", option)
                    print(f"✅ 종료 월 {end_month} 선택 완료.")
                    found_end_month = True
                    break

            if not found_end_month:
                print(f"❌ 종료 월 {end_month} 찾을 수 없음!")

            # ✅ 종료 일 선택
            print(f"🔍 종료 일 {end_day} 선택 중...")
            end_day_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "(//div[@class='select w3'])[4]//span[@class='select_btn']")))
            driver.execute_script("arguments[0].click();", end_day_btn)
            time.sleep(1)

            options = driver.find_elements(By.XPATH, "//div[@class='select w3']//ul[@class='select_list scroll_cst']//li/a")
            for option in options:
                if option.text.strip() == end_day:
                    driver.execute_script("arguments[0].click();", option)
                    print(f"✅ 종료 일 {end_day} 선택 완료.")
                    found_end_day = True
                    break

            if not found_end_day:
                print(f"❌ 종료 일 {end_day} 찾을 수 없음!")

            print("🔍 시작일 선택 시작...")

            # ✅ 시작 연도 선택
            start_year_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "(//div[@class='select w2'])[1]//span[@class='select_btn']")))
            driver.execute_script("arguments[0].click();", start_year_btn)
            time.sleep(1)

            options = driver.find_elements(By.XPATH, "//div[@class='select w2']//ul[@class='select_list scroll_cst']//li/a")
            for option in options:
                if option.text.strip() == start_year:
                    driver.execute_script("arguments[0].click();", option)
                    print(f"✅ 시작 연도 {start_year} 선택 완료.")
                    found_start_year = True
                    break

            if not found_start_year:
                print(f"❌ 시작 연도 {start_year} 찾을 수 없음!")

            # ✅ 시작 월 선택
            start_month_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "(//div[@class='select w3'])[1]//span[@class='select_btn']")))
            driver.execute_script("arguments[0].click();", start_month_btn)
            time.sleep(1)

            options = driver.find_elements(By.XPATH, "//div[@class='select w3']//ul[@class='select_list scroll_cst']//li/a")
            for option in options:
                if option.text.strip() == start_month:
                    driver.execute_script("arguments[0].click();", option)
                    print(f"✅ 시작 월 {start_month} 선택 완료.")
                    found_start_month = True
                    break

            if not found_start_month:
                print(f"❌ 시작 월 {start_month} 찾을 수 없음!")

            # ✅ 시작일 선택
            print(f"🔍 시작 일 {start_day} 선택 중...")
            start_day_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "(//div[@class='select w3'])[2]//span[@class='select_btn']")))
            driver.execute_script("arguments[0].click();", start_day_btn)
            time.sleep(1)

            options = driver.find_elements(By.XPATH, "//div[@class='select w3']//ul[@class='select_list scroll_cst']//li/a")
            for option in options:
                if option.text.strip() == start_day:
                    driver.execute_script("arguments[0].click();", option)
                    print(f"✅ 시작 일 {start_day} 선택 완료.")
                    found_start_day = True
                    break

            if not found_start_day:
                print(f"❌ 시작 일 {start_day} 찾을 수 없음!")

            # ❌ 종료일 또는 시작일을 찾지 못한 경우, 다시 선택하도록 처리
            if not (found_end_year and found_end_month and found_end_day):
                print(f"⚠️ 종료일 ({end_year}-{end_month}-{end_day}) 선택 실패. 다시 시도합니다...")
                retry_count += 1
                continue

            if not (found_start_year and found_start_month and found_start_day):
                print(f"⚠️ 시작일 ({start_year}-{start_month}-{start_day}) 선택 실패. 다시 시도합니다...")
                retry_count += 1
                continue

            # ✅ 최종적으로 모든 날짜 선택 성공
            period = f"{start_year}-{start_month}-{start_day} ~ {end_year}-{end_month}-{end_day}"
            print(f"✅ 최종 기간 선택 완료: {period}")
            break  # while 루프 종료

        except Exception as e:
            print(f"❌ 오류 발생: {e}")
            retry_count += 1
            continue  # 오류 발생 시 다시 실행

    if retry_count == 3:
        print("🚨 3회 이상 재시도 실패. 기간 선택을 중단합니다.")



try:
    # 1. 네이버 데이터랩 접속
    url = 'https://datalab.naver.com/shoppingInsight/sCategory.naver'
    driver.get(url)
    time.sleep(3)

    # 2. '생활/건강' 카테고리 버튼 클릭 (자바스크립트 사용)
    category_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@class='select_btn']")))
    driver.execute_script("arguments[0].click();", category_btn)
    time.sleep(1)

    # 3. 드롭다운 메뉴에서 '생활/건강' 카테고리 선택
    category_life_health = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@data-cid='50000008']")))
    category_life_health.click()
    time.sleep(2)
    category_1 = "생활_건강"
    print("✅ '생활/건강' 카테고리 선택 완료!")

    
    subcategory_btn = wait.until(EC.presence_of_element_located((By.XPATH, "(//span[@class='select_btn'])[2]")))
    driver.execute_script("arguments[0].click();", subcategory_btn)
    category_2 = "자동차용품"
    time.sleep(1)

    subcategory_car = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@data-cid='50000055']")))
    subcategory_car.click()
    time.sleep(2)
    
    print("✅ '자동차용품' 2차 카테고리 선택 완료!")
    # 5. 기기별 선택
    device_all_checkbox = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//input[@id='18_device_0']")))  # 기기별 > 전체 체크박스
    driver.execute_script("arguments[0].click();", device_all_checkbox)
    time.sleep(1)
    device_selection = "전체"
    print("✅ '기기별 > 전체' 선택 완료")

    # 6. 성별 선택
    sex_all_checkbox = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//input[@id='19_gender_0']")))  # 성별 > 전체 체크박스
    driver.execute_script("arguments[0].click();", sex_all_checkbox)
    time.sleep(1)
    gender_selection = "전체"
    print("✅ '성별 > 전체' 선택 완료")

    # 7. 연령 선택
    age_all_checkbox = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//input[@id='20_age_0']")))  # 연령 > 전체 체크박스
    driver.execute_script("arguments[0].click();", age_all_checkbox)
    time.sleep(1)
    age_selection = "전체"
    print("✅ '연령 > 전체' 선택 완료")

    
    # 1. for 루프: 옵션 클릭 (50000935 제외)
    for i in range(0, 16):
        option_id = 50000948 - i
        if option_id == 50000935:
            continue

        # data-cid 값으로 해당 <a> 요소 찾기
        xpath_query = f"//a[@data-cid='{option_id}']"
        
        try:
            option_element = driver.find_element(By.XPATH, xpath_query)
            driver.execute_script("arguments[0].click();", option_element)
            time.sleep(0.5)
        except Exception as e:
            print(f"Element with data-cid {option_id} not found: {e}")

        # 2. for 루프 종료 후 3차 카테고리 최종 선택
        subcategory_3_btn = wait.until(EC.presence_of_element_located((By.XPATH, "(//span[@class='select_btn'])[3]")))
        driver.execute_script("arguments[0].click();", subcategory_3_btn)
        time.sleep(1)
        print("✅3차 카테고리 선택 완료!")
        
        category_3_mapping = {
            "50000933": "DIY용품",
            "50000934": "램프",
            "50000936": "배터리용품",
            "50000937": "공기청정용품",
            "50000938": "세차용품",
            "50000939": "키용품",
            "50000940": "편의용품",
            "50000941": "오일-소모품",
            "50000942": "익스테리어용품",
            "50000943": "인테리어용품",
            "50000944": "전기용품",
            "50000945": "수납용품",
            "50000946": "휴대폰용품",
            "50000947": "타이어-휠",
            "50000948": "튜닝용품"
        }
        def get_category_name(option_id):
            option_id = str(option_id).strip()  # 문자열 변환 + 공백 제거
            return category_3_mapping.get(option_id, "알 수 없는 카테고리") 
        category_3 = get_category_name(option_id)


        try:
            current_date = datetime(2024, 7, 1)
            final_date = datetime(2024, 8, 4)
            week_counter = 1

            while current_date <= final_date:
                week_start = current_date
                week_end = week_start + timedelta(days=6)
                if week_end > final_date:
                    week_end = final_date
                    

                start_year, start_month, start_day = str(week_start.date()).split("-")
                end_year, end_month, end_day = str(week_end.date()).split("-")
                
                select_period(start_year=start_year, start_month=start_month, start_day=start_day,
                            end_year=end_year, end_month=end_month, end_day=end_day)
                
                print(f"✅ {week_counter}주차 기간 선택 완료: {week_start.date()} ~ {week_end.date()}")
                
                current_date = week_end + timedelta(days=1)
                week_counter += 1

                
                # 8. 조회하기 버튼 클릭
                search_button = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//a[@class='btn_submit']/span[text()='조회하기']")))  # 조회하기 버튼
                driver.execute_script("arguments[0].click();", search_button)
                time.sleep(2)
                print("✅ 조회버튼 클릭 완료!")
                
                # 9. 결과 데이터 추출
                
                time.sleep(3)

                download_folder = os.path.expanduser("~/Downloads")  # Windows/Linux/Mac 지원

                target_folder = "./downloaded_data"

                    # 🔹 대상 폴더가 없으면 생성
                if not os.path.exists(target_folder):
                        os.makedirs(target_folder)

                try:
                            # ✅ "조회결과 다운로드" 버튼 클릭
                        download_button = wait.until(EC.element_to_be_clickable(
                                (By.XPATH, "//a[@class='btn_document_down' and text()='조회결과 다운로드']")))
                        driver.execute_script("arguments[0].click();", download_button)
                        print("✅ '조회결과 다운로드' 버튼 클릭 완료!")

                        time.sleep(5)  # 네트워크 속도에 따라 조절 가능

                            # ✅ 최근 다운로드된 파일 찾기 (파일명 패턴이 필요할 경우 수정 가능)
                        files = os.listdir(download_folder)
                        files = [f for f in files if f.endswith(".csv")]

                        new_file_name = os.path.join(target_folder, f"{week_start.date()} ~ {week_end.date()}_{category_3}.xlsx")

                        if not files:
                                print("❌ 다운로드된 파일을 찾을 수 없습니다.")
                        else:
                            latest_file = max([os.path.join(download_folder, f) for f in files], key=os.path.getctime)
                            latest_file_df = pd.read_csv(latest_file, skiprows=7, header=0, encoding="utf-8")
                            with pd.ExcelWriter(new_file_name, engine="xlsxwriter") as writer:
                                latest_file_df.to_excel(writer, sheet_name="Sheet1", index=False)
                                worksheet = writer.sheets["Sheet1"]
                                for col_idx, col_name in enumerate(latest_file_df.columns):
                                    max_len = latest_file_df[col_name].astype(str).map(len).max()
                                    max_len = max(max_len, len(col_name)) + 2
                                    worksheet.set_column(col_idx, col_idx, max_len)

                                # ✅ 파일 이동 및 이름 변경
                            print(f"✅ 다운로드된 파일 이동 및 이름 변경 완료")

                except Exception as e:
                        print(f"❌ 데이터 추출 오류: {e}")
                    
                # 10. 기기별 / 성별 / 연령별 비중 추출
                try:
                    time.sleep(3)
                    buttons = driver.find_elements(By.CLASS_NAME, "btn_trend_view")
                    print(f"🔍 찾은 버튼 수: {len(buttons)}개")
                    wait = WebDriverWait(driver, 10)
                    driver.execute_script("window.scrollBy(0, -100);")
                    
                    def close_popup():
                        try:
                            close_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn_popup_close")))
                            close_btn.click()
                            print("🔒 닫기 버튼 클릭 완료\n")
                            time.sleep(2)
                        except Exception as e:
                            print(f"⚠️ 닫기 버튼 클릭 실패: {e}")
                    
                    if len(buttons) >= 1:
                        print("📊 [기기별 비중] 버튼 클릭 중...")
                        driver.execute_script("arguments[0].scrollIntoView(true);", buttons[0])
                        buttons[0].click()
                        time.sleep(2)
                        device_data = driver.find_element(By.CSS_SELECTOR, ".pie_chart").text
                        print(f"✅ 기기별 비중:\n{device_data}\n")
                        close_popup()
                    if len(buttons) >= 2:
                        print("📊 [성별 비중] 버튼 클릭 중...")
                        driver.execute_script("arguments[0].scrollIntoView(true);", buttons[1])
                        buttons[1].click()
                        time.sleep(2)
                        gender_data = driver.find_element(By.CSS_SELECTOR, ".pie_chart").text
                        print(f"✅ 성별 비중:\n{gender_data}\n")
                        close_popup()
                    if len(buttons) >= 3:
                        print("📊 [연령별 비중] 버튼 클릭 중...")
                        driver.execute_script("arguments[0].scrollIntoView(true);", buttons[2])
                        buttons[2].click()
                        time.sleep(2)
                        age_data = driver.find_element(By.CSS_SELECTOR, ".pie_chart").text
                        print(f"✅ 연령별 비중:\n{age_data}\n")
                        close_popup()
                except Exception as e:
                    print(f"❌ 추출 오류: {e}")
                    
                # 11. 인기검색어 추출
                data = {"인기검색어 순위": [], "인기검색어": []}

                try:
                    print("🔍 [인기검색어 500개] 추출 중...")
                    popular_keywords = []
                    while True:
                        keywords_elements = driver.find_elements(By.CSS_SELECTOR, ".rank_top1000_list li a")
                        for element in keywords_elements:
                            rank = element.find_element(By.CLASS_NAME, "rank_top1000_num").text
                            keyword = element.text.replace(rank, "").strip()
                            popular_keywords.append((rank, keyword))
                            if len(popular_keywords) >= 500:
                                print("🎯 500개 키워드 추출 완료!\n")
                                break
                        print(f"📜 현재까지 추출된 키워드 수: {len(popular_keywords)}개")
                        if len(popular_keywords) >= 500:
                            break
                        try:
                            next_button = driver.find_element(By.CSS_SELECTOR, ".btn_page_next")
                            if "disabled" in next_button.get_attribute("class") or not next_button.is_enabled():
                                print("✅ 마지막 페이지 도달. 추출 완료!\n")
                                break
                            else:
                                print("➡️ 다음 페이지로 이동 중...")
                                driver.execute_script("arguments[0].click();", next_button)
                                time.sleep(2)
                        except Exception as e:
                            print(f"⚠️ 다음 페이지 버튼 클릭 실패 또는 마지막 페이지 도달: {e}")
                            break
                    print(f"\n🎯 인기검색어 총 {len(popular_keywords)}개 추출 완료!\n")
                    for rank, keyword in popular_keywords:
                        data["인기검색어 순위"].append(rank)
                        data["인기검색어"].append(keyword)
                        print(f"{rank}. {keyword}")
                except Exception as e:
                    print(f"❌ 인기검색어 데이터 추출 오류: {e}")
                    
                df = pd.DataFrame(data)
                output_file = f"./crawled_data/{period}_{category_1}_{category_2}_{category_3}_{device_selection}_{gender_selection}_{age_selection}.xlsx"
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='조회결과')
                print(f"엑셀 파일 '{output_file}'로 저장 완료!")

                #엑셀 파일 병합
                date_range_str = f"{week_start.date()} ~ {week_end.date()}"

                excel_folder_1 = "./downloaded_data"
                excel_folder_2 = "./crawled_data"

                files_1 = [os.path.join(excel_folder_1, f) for f in os.listdir(excel_folder_1)
                    if f.endswith(".xlsx") and date_range_str in f]
                files_2 = [os.path.join(excel_folder_2, f) for f in os.listdir(excel_folder_2)
                    if f.endswith(".xlsx") and date_range_str in f]

                df_list_1 = [pd.read_excel(file) for file in files_1]
                df_list_2 = [pd.read_excel(file) for file in files_2]

                merged_df_1 = pd.concat(df_list_1, ignore_index=True) if df_list_1 else pd.DataFrame()
                merged_df_2 = pd.concat(df_list_2, ignore_index=True) if df_list_2 else pd.DataFrame()

                target_folder_1 = "./merged_data"
                merged_file_path = os.path.join(target_folder_1, f"{week_counter - 1}주차_{category_3}.xlsx")

                with pd.ExcelWriter(merged_file_path, engine="xlsxwriter") as writer:
                    merged_df_1.to_excel(writer, sheet_name="조회수 총합", index=False)
                    merged_df_2.to_excel(writer, sheet_name="인기검색어 TOP500", index=False)

                print(f"병합된 엑셀 파일 생성 완료")

                folders_to_clean = [
                    "./crawled_data",
                    "./downloaded_data"
                ]               

                for folder in folders_to_clean:
                    for filename in os.listdir(folder):
                        file_path = os.path.join(folder, filename)
                        try:
                            if os.path.isfile(file_path) or os.path.islink(file_path):
                                os.remove(file_path)
                            elif os.path.isdir(file_path):
                                shutil.rmtree(file_path)
                        except Exception as e:
                            print(f"삭제 실패: {file_path} - {e}")

                print("파일 정리 완료")


                
        except Exception as e:
            print(f"❌ 오류 발생: {e}")
            

except Exception as e:
    print(f"❌ 오류 발생: {e}")
            
finally:
    driver.quit()
    
 
