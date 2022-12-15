import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime
import os


def search_keywords(browser, urls, keywords, pages):
    # Create a new dataframe to store the results
    results = pd.DataFrame(columns=["Link", "Title", "Keyword"])

    # Iterate through the keywords and URLs
    for url, keyword, page in zip(urls, keywords, pages):
        # Navigate to the URL
        browser.get(url)

        time.sleep(2)
        # Find the consent button and click on it
        try:
            consent_button = browser.find_element(By.ID, "L2AGLb")
            if consent_button.is_displayed():
                # Consent button is visible, so click on it
                consent_button.click()
            time.sleep(1)
        except NoSuchElementException:
            pass

        # Find the search input box and enter the keyword
        search_input = browser.find_element("name", "q")
        search_input.send_keys(keyword)

        # Submit the search form
        search_input.send_keys(Keys.RETURN)

        # Wait for the results to load
        # You may need to adjust this value depending on your internet speed
        browser.implicitly_wait(3)

        # Iterate through the pages of search results
        for _ in range(page):
            # Find the "Next" button and click on it
            next_button = browser.find_element(By.ID, "pnnext")
            next_button.click()

            # Wait for the results to load
            # You may need to adjust this value depending on your internet speed
            time.sleep(1)

        # Find all the result elements on the page
        result_elements = browser.find_elements(By.CSS_SELECTOR, ".yuRUbf")

        # Extract the text from each result element and add it to the dataframe
        for result in result_elements:
            title = result.find_element(By.CSS_SELECTOR, "h3.DKV0Md").text
            link = result.find_element(By.CSS_SELECTOR, "a").get_attribute('href')
            link = link.format(link)
            data = {"Keyword": keyword, "Title": title, "Link": link}
            temp_df = pd.DataFrame(data, index=range(len(data)))
            results = pd.concat([results, temp_df], ignore_index=True)

        # results["Link"] = results.apply(lambda row: "=HYPERLINK(\"{}\")".format(row["Link"]), axis=1)

    return results


def start_browser():
    # Start a new Chrome browser, add binary location to avoid WebDriverException
    options = webdriver.ChromeOptions()
    options.binary_location = r"C:\Program Files\Google\Chrome Beta\Application\chrome.exe"
    service_obj = Service(r"C:\Users\a.sobczyk\Desktop\Data_Science\webdrivers\chromedriver.exe")
    browser = webdriver.Chrome(service=service_obj, options=options)
    browser.minimize_window()

    return browser


def close_browser(browser):
    # Close the browser
    browser.close()


def read_keywords_and_urls(file_name):
    # Read the list of keywords and URLs from an Excel file
    df = pd.read_excel(file_name)

    return df["Link"], df["Keyword"], df["Page"]


def save_results(results, file_name):
    # Set the output folder path
    output_folder = "output/"

    # Check of folder exists, if no then create new one
    if not os.path.isdir(output_folder):
        os.mkdir(output_folder)

    # Get current date and time and format it as a string
    curent_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    # Save the results dataframe to an Excel file in the output folder
    file_name = file_name + "_" + curent_datetime + ".xlsx"
    results.to_excel(os.path.join(output_folder, file_name), index=False)


# UCOMMENT THIS FUNCTION TO RUN IT WITHOUT GUI

# def main():
#     # Start the browser
#     browser = start_browser()
#
#     # Read the keywords and URLs
#     urls, keywords, pages = read_keywords_and_urls("links.xlsx")
#
#     # Search for the keywords and get the results
#     results = search_keywords(browser, urls, keywords, pages)
#
#     # Get date and time
#     now = datetime.now()
#
#     # Create filename
#     file_name = "results_" + now.strftime("%Y_%m_%d_%H-%M-%S") + ".xlsx"
#
#     # Save the results
#     save_results(results, file_name)
#     print(results)
#
#     # Close the browser
#     close_browser(browser)
#
#
# if __name__ == "__main__":
#     main()
