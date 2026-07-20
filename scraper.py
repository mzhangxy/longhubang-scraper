from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

def run_scraper():
    print("初始化浏览器环境...")
    options = webdriver.ChromeOptions()
    options.add_argument('--headless') # 无头模式，不打开实体浏览器界面
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36')

    driver = webdriver.Chrome(options=options)
    driver.get("https://data.10jqka.com.cn/market/lhbyyb/orgcode/ff5f15dad72400af/")

    all_data = []
    total_pages = 67

    for page in range(1, total_pages + 1):
        print(f"正在抓取第 {page} 页...")
        try:
            # 等待表格加载出来（最多等 10 秒）
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table.m-table tbody tr"))
            )
            time.sleep(1.5) # 给一点额外渲染时间

            # 解析表格行
            rows = driver.find_elements(By.CSS_SELECTOR, "table.m-table tbody tr")
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 8:
                    all_data.append({
                        "上榜日期": cols[0].text.strip(),
                        "股票简称": cols[1].text.strip(),
                        "上榜原因": cols[2].text.strip(),
                        "涨跌幅(%)": cols[3].text.strip(),
                        "买入额（万）": cols[4].text.strip(),
                        "卖出额（万）": cols[5].text.strip(),
                        "买卖净额（万）": cols[6].text.strip(),
                        "所属板块": cols[7].text.strip()
                    })

            # 如果不是最后一页，则点击“下一页”
            if page < total_pages:
                next_btn = driver.find_element(By.LINK_TEXT, "下一页")
                driver.execute_script("arguments[0].click();", next_btn) # 使用 JS 点击防止被遮挡
                time.sleep(2) # 等待页面数据刷新
        except Exception as e:
            print(f"第 {page} 页抓取出现问题，可能会跳过此页。报错信息: {e}")

    driver.quit()

    # 保存为 Excel
    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel("龙虎榜数据.xlsx", index=False)
        print(f"抓取完成，共获取 {len(all_data)} 条数据，已保存为 龙虎榜数据.xlsx")
    else:
        print("未能抓取到任何数据，请检查网页结构是否变化。")

if __name__ == "__main__":
    run_scraper()
