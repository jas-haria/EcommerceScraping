from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials
import time
import numpy as np


filenames = ["earbuds.xlsx", "headphones.xlsx", "speakers.xlsx", "wired earphones.xlsx"]
file_ids = ["1Hn7wEEtDfJMhFsqfnVBctsDHtqM6ZdDz", "1DtfJS6bGyYw89Jd62dotrgXqDlmP79NI",
            "18z3vERjAzzG3gJoSyTxSqS4gF4w0dfnT", "1kwDAkOUxAs6e6D8AyTDRMqiM5rKfTupd"]
weekly_sales_column = "Weekly Sales (Ratings * 5)"
total_sales_column = "Total Sales (Total Ratings * 5)"
weekly_by_total_sales_column = "Weekly Sales / Total Sales"
sleep_time = 10


def get_driver():
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver


def google_auth():
    gauth = GoogleAuth()
    scope = ["https://www.googleapis.com/auth/drive"]
    gauth.credentials = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    return GoogleDrive(gauth)


def scrape_amazon(driver, df_amazon_urls, current_date, df_amazon_raw):
    ratings = list()
    reviews = list()
    rating = None
    review = None
    for url in df_amazon_urls['URL']:
        try:
            driver.get(url)
            time.sleep(sleep_time)
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            raw_data = soup.find('div', attrs={'data-hook': 'cr-filter-info-review-rating-count'}).text.strip()
            split_data = raw_data.split("total ratings,")
            rating = int(split_data[0].strip().replace(",", ""))
            review = int(split_data[1].split("with reviews")[0].strip().replace(",", ""))
        except:
            rating = np.nan
            review = np.nan
        finally:
            ratings.append(rating)
            reviews.append(review)
    df_amazon_raw = df_amazon_raw.drop(columns="Rat - " + current_date, axis=1, errors='ignore')
    df_amazon_raw = df_amazon_raw.drop(columns="Rev - " + current_date, axis=1, errors='ignore')
    if len(df_amazon_raw) < len(ratings):
        difference = len(ratings) - len(df_amazon_raw)
        columns_length = len(df_amazon_raw.columns)
        dummy_list = [np.nan] * columns_length
        for i in range(difference):
            df_amazon_raw = df_amazon_raw.append(pd.DataFrame([dummy_list],
                                                              columns=list(df_amazon_raw.columns)),
                                                 ignore_index=True)
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
            time.sleep(sleep_time)
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            raw_rating_data = soup.find(lambda tag: tag.name == "span" and "Ratings" in tag.text and "Reviews" not in tag.text).get_text(strip=True)
            split_data = raw_rating_data.split("Ratings")
            rating = int(split_data[0].strip().replace(",", ""))
            raw_review_data = soup.find(lambda tag: tag.name == "span" and "Ratings" not in tag.text and "Reviews" in tag.text).get_text(strip=True)
            split_data = raw_review_data.split("Reviews")
            review = int(split_data[0].strip().replace(",", ""))
        except:
            rating = np.nan
            review = np.nan
        finally:
            ratings.append(rating)
            reviews.append(review)

    df_flipkart_raw = df_flipkart_raw.drop(columns="Rat - " + current_date, axis=1, errors='ignore')
    df_flipkart_raw = df_flipkart_raw.drop(columns="Rev - " + current_date, axis=1, errors='ignore')
    if len(df_flipkart_raw) < len(ratings):
        difference = len(ratings) - len(df_flipkart_raw)
        columns_length = len(df_flipkart_raw.columns)
        dummy_list = [np.nan] * columns_length
        for i in range(difference):
            df_flipkart_raw = df_flipkart_raw.append(pd.DataFrame([dummy_list],
                                                                  columns=list(df_flipkart_raw.columns)),
                                                     ignore_index=True)
    df_flipkart_raw.insert(loc=0, column="Rat - " + current_date, value=ratings)
    df_flipkart_raw.insert(loc=0, column="Rev - " + current_date, value=reviews)
    return df_flipkart_raw


def get_daily(df, df_urls):
    num_of_cols = len(df.columns)
    for i in range(0, num_of_cols - 2):
        df.iloc[:, i] = df.iloc[:, i] - df.iloc[:, i + 2]
    columns_to_delete = [x for x in df.columns if str(x).startswith("Rev")]
    df = df.drop(columns=columns_to_delete, axis=1)
    df_urls[weekly_sales_column] = 0
    df_urls[total_sales_column] = 0
    df_urls[weekly_by_total_sales_column] = 0
    for i in range(0, len(df)):
        df_urls[weekly_sales_column][i] = 0
        df_urls[total_sales_column][i] = 0
        df_urls[weekly_by_total_sales_column][i] = 0
        for j in range(0, len(df.columns)):
            if j < 7:
                df_urls[weekly_sales_column][i] += df.iloc[i, j] if not np.isnan(df.iloc[i, j]) else 0
            df_urls[total_sales_column][i] += df.iloc[i, j] if not np.isnan(df.iloc[i, j]) else 0
        df_urls[weekly_sales_column][i] *= 5
        df_urls[total_sales_column][i] *= 5
        if df_urls[total_sales_column][i] > 0:
            weekly_by_total_sales = df_urls[weekly_sales_column][i] * 100 / df_urls[total_sales_column][i]
            df_urls[weekly_by_total_sales_column][i] = round(weekly_by_total_sales, 2)
    df = df_urls.join(df)
    return df


def scrape_data(filename, file_id):
    drive = google_auth()
    file = drive.CreateFile({'id': file_id})
    file.GetContentFile(filename)
    df_amazon_urls = pd.read_excel(filename, sheet_name="Amazon URL")
    df_flipkart_urls = pd.read_excel(filename, sheet_name="Flipkart URL")
    df_amazon_raw = pd.read_excel(filename, sheet_name="Amazon - Raw")
    df_flipkart_raw = pd.read_excel(filename, sheet_name="Flipkart - Raw")
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    driver = get_driver()
    current_date = str(datetime.now().day) + "/" + str(datetime.now().month)

    df_amazon_urls.to_excel(excel_writer=writer, index=False, sheet_name="Amazon URL")
    df_flipkart_urls.to_excel(excel_writer=writer, index=False, sheet_name="Flipkart URL")

    df_amazon_raw = scrape_amazon(driver, df_amazon_urls, current_date, df_amazon_raw)
    df_flipkart_raw = scrape_flipkart(driver, df_flipkart_urls, current_date, df_flipkart_raw)

    df_amazon_raw.to_excel(excel_writer=writer, index=False, sheet_name="Amazon - Raw")
    df_flipkart_raw.to_excel(excel_writer=writer, index=False, sheet_name="Flipkart - Raw")

    df_flipkart_daily = get_daily(df_flipkart_raw, df_flipkart_urls)
    df_amazon_daily = get_daily(df_amazon_raw, df_amazon_urls)

    df_amazon_daily.to_excel(excel_writer=writer, index=False, sheet_name="Amazon - Daily")
    df_flipkart_daily.to_excel(excel_writer=writer, index=False, sheet_name="Flipkart - Daily")

    driver.close()
    writer.close()
    update_file = drive.CreateFile({'id': file_id})
    update_file.SetContentFile(filename)
    update_file.Upload()


for f_name, f_id in zip(filenames, file_ids):
    scrape_data(f_name, f_id)
