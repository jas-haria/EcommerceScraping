from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime


filename = "data.xlsx"


def get_driver():
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver


def scrape_amazon(driver, df_amazon_urls, current_date, df_amazon_raw):
    ratings = list()
    reviews = list()
    rating = None
    review = None
    for url in df_amazon_urls['URL']:
        try:
            driver.get(url)
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            raw_data = soup.find('div', attrs={'data-hook': 'cr-filter-info-review-rating-count'}).text.strip()
            split_data = raw_data.split("total ratings,")
            rating = split_data[0].strip().replace(",", "")
            review = split_data[1].split("with reviews")[0].strip().replace(",", "")
        except:
            rating = -1
            review = -1
        finally:
            ratings.append(rating)
            reviews.append(review)
    df_amazon_raw = df_amazon_raw.drop(columns="Rat - " + current_date, axis=1, errors='ignore')
    df_amazon_raw = df_amazon_raw.drop(columns="Rev - " + current_date, axis=1, errors='ignore')
    df_amazon_raw.insert(loc=0, column="Rat - " + current_date, value=ratings)
    df_amazon_raw.insert(loc=0, column="Rev - " + current_date, value=reviews)
    return df_amazon_raw


def scrape_flipkart(driver, df_flipkart_urls, current_date, df_flipkart_raw):
    ratings = list()
    reviews = list()
    rating = None
    review = None
    for url in df_flipkart_urls['URL']:
        try:
            driver.get(url)
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            raw_rating_data = soup.find(lambda tag: tag.name == "span" and "Ratings" in tag.text and "Reviews" not in tag.text).get_text(strip=True)
            split_data = raw_rating_data.split("Ratings")
            rating = split_data[0].strip().replace(",", "")
            raw_review_data = soup.find(lambda tag: tag.name == "span" and "Ratings" not in tag.text and "Reviews" in tag.text).get_text(strip=True)

            split_data = raw_review_data.split("Reviews")
            review = split_data[0].strip().replace(",", "")
        except:
            rating = -1
            review = -1
        finally:
            ratings.append(rating)
            reviews.append(review)

    df_flipkart_raw = df_flipkart_raw.drop(columns="Rat - " + current_date, axis=1, errors='ignore')
    df_flipkart_raw = df_flipkart_raw.drop(columns="Rev - " + current_date, axis=1, errors='ignore')
    df_flipkart_raw.insert(loc=0, column="Rat - " + current_date, value=ratings)
    df_flipkart_raw.insert(loc=0, column="Rev - " + current_date, value=reviews)
    return df_flipkart_raw


def scrape_data():
    df_amazon_urls = pd.read_excel(filename, sheet_name="Amazon URL")
    df_flipkart_urls = pd.read_excel(filename, sheet_name="Flipkart URL")
    df_amazon_raw = pd.read_excel(filename, sheet_name="Amazon - Raw")
    df_flipkart_raw = pd.read_excel(filename, sheet_name="Amazon - Raw")
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    driver = get_driver()
    current_date = str(datetime.now().day) + "/" + str(datetime.now().month)

    df_amazon_raw = scrape_amazon(driver, df_amazon_urls, current_date, df_amazon_raw)
    df_flipkart_raw = scrape_flipkart(driver, df_flipkart_urls, current_date, df_flipkart_raw)

    df_amazon_urls.to_excel(excel_writer=writer, index=False, sheet_name="Amazon URL")
    df_flipkart_urls.to_excel(excel_writer=writer, index=False, sheet_name="Flipkart URL")
    df_amazon_raw.to_excel(excel_writer=writer, index=False, sheet_name="Amazon - Raw")
    df_flipkart_raw.to_excel(excel_writer=writer, index=False, sheet_name="Flipkart - Raw")
    driver.close()
    writer.close()


scrape_data()
