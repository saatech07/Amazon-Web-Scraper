import time
import pandas as pd
from selenium import webdriver
from spire.xls import *
from spire.xls.common import *
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from amazoncaptcha import AmazonCaptcha


def solve_captcha(image_link):
    captcha = AmazonCaptcha.fromlink(image_link)
    solution = captcha.solve()
    return solution

def scrape_data_from_amazon(search_msg):
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get("https://www.amazon.com/")
    captcha_solver = driver.find_element(By.XPATH, "//div[@class='a-box']/div/div/img[@src]")
    captcha_src = captcha_solver.get_attribute('src')
    captcha_text = solve_captcha(captcha_src)
    solving_captcha = driver.find_element(By.XPATH, "//input[@placeholder = 'Type characters']")
    solving_captcha.click()
    solving_captcha.send_keys(captcha_text)
    solving_captcha.submit()
    search_bar = driver.find_element(By.XPATH, "//div[@class='nav-fill']//input[@id='twotabsearchtextbox']")
    search_bar.click()
    search_bar.send_keys(search_msg)
    search_bar.submit()
    page_range = driver.find_element(By.XPATH,
                                     "//span[@class='s-pagination-strip']//span[@class='s-pagination-item s-pagination-disabled']")
    product_data = []
    price_symbol = driver.find_element(By.XPATH, "//span[@class='a-price']/span[2]/span[@class='a-price-symbol']")
    txt = price_symbol.text
    for i in range(int(page_range.text)):
        time.sleep(4)
        all_products = driver.find_elements(By.XPATH, '//div[starts-with(@class,"s-widget-container s-spacing-small s-widget-container-height-small celwidget slot=MAIN template=SEARCH_RESULTS widgetId=search-results")]')
        titles = driver.find_elements(By.XPATH, "//div[@data-component-type='s-search-result']//h2//a//span")
        price_list = []
        sold_list = []
        global_rating_list = []
        all_ratings = []
        for item in all_products:
            try:
                price_element = item.find_element(By.XPATH, './/span[@class="a-price-whole"]')
                price_list.append(price_element.text)
            except Exception as e:
                price_list.append('None')

            try:
                sold_element = item.find_element(By.XPATH, './/span[contains(text(),"bought")]')
                sold_list.append(sold_element.text)
            except Exception as e:
                sold_list.append('None')

            try:
                global_rated_element = item.find_element(By.XPATH, './/span[@class="a-size-base s-underline-text"]')
                global_rating_list.append(global_rated_element.text)
            except Exception as e:
                global_rating_list.append('None')

            try:
                total_rating_element = item.find_element(By.XPATH, './/span[contains(@aria-label, "stars")]')
                all_ratings.append(total_rating_element.get_attribute('aria-label'))
            except Exception as e:
                all_ratings.append('None')

        # image section
        images = driver.find_elements(By.XPATH,
                                      "//div[@data-component-type='s-search-result']//img[@data-image-latency='s-product-image']")
        all_images = []
        for imgs in images:
            image = imgs.get_attribute('src')
            all_images.append(image)
        for j in range(len(titles)):
            temp_data = {'Image:': all_images[j],
                         'Title:': titles[j].text,
                         'No of Products Sold:': sold_list[j],
                         'Global Ratings:': global_rating_list[j],
                         'Rating out of 5:': all_ratings[j],
                         'Price:': f'{txt} {price_list[j]}'
                         }
            product_data.append(temp_data)
            if len(all_products) == j+1:
                try:
                    next_page = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[starts-with(@aria-label, 'Go to next page')]")))
                    next_page.click()
                except Exception:
                    break
        df_data = pd.DataFrame(product_data)
        df_data.to_excel('Laptop_scraped_data.xlsx', index=False)
        driver.close()
if __name__ == "__main__":

    msg = input("Enter the keyword or name you want to scrap data for: ").lower()
    scrape_data_from_amazon(msg)
