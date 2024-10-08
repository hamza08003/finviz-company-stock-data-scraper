# Import all necessary libraries
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import pandas as pd
import os
import time


driver = webdriver.Chrome()
wait = WebDriverWait(driver, 10)

print("Opening Finviz .....")
driver.get('https://finviz.com/screener.ashx')


total_data_tables_legth = 0


def apply_filters_and_conditions():
    # Locate the custom_button_element using its XPath and click on it
    custom_button_element = wait.until(EC.element_to_be_clickable((By.XPATH, '//td[@class="screener-view-button"]/a[text()="Custom"]')))
    custom_button_element.click()

    time.sleep(3)

    # Locate all the radio elements and click on them to uncheck them
    radio_buttons_to_toggle = [[2, 1], [2, 5],
                               [2, 6], [2, 7], [2, 8], [2, 9], [8, 7]]

    print("Checking/Uncheking Required Coloumns .....")
    # Unchecking/Checking
    for i in radio_buttons_to_toggle:
        element = driver.find_element(By.XPATH, f'//table[@class="screener-groups_settings-table is-custom"]/tbody/tr[{i[0]}]/td[{i[1]}]')
        element.click()
        time.sleep(3)

    filters = driver.find_element(
        By.CSS_SELECTOR, 'a[href="screener.ashx?v=151"]')
    filters.click()

    time.sleep(3)

    print("Applying Relative and Current Volume Filters .....")
    relative_vol = driver.find_element(By.ID, "fs_sh_relvol")
    Select(relative_vol).select_by_index(1)

    time.sleep(3)

    relative_vol = driver.find_element(By.ID, "fs_sh_curvol")
    Select(relative_vol).select_by_index(14)

    time.sleep(3)


def find_total_data_tables():
    global total_data_tables_legth

    # Finding the <td> element with the ID 'screener_pagination'
    td_element = driver.find_element(By.ID, "screener_pagination")

    print("Finding total number of data tables .....")
    # Finding all <a> tags within the <td> element
    a_tags = td_element.find_elements(By.TAG_NAME, "a")

    # Iterating through each <a> tag until the one with class 'is-next' is encountered
    for a_tag in a_tags:
        # Exclude the <a> tag with class 'is-next'
        if 'is-next' not in a_tag.get_attribute('class'):
            # Add the length of the current <a> tag to the total length
            total_data_tables_legth += len(a_tag.text)


def fetch_and_save_data_tables_html_content():
    global total_data_tables_legth

    tabele_no = 1

    print("Fetching and Saving all Data Tables' HTML Content in a txt file .....")
    # Find the table element
    table_element = driver.find_element(By.CLASS_NAME, 'styled-table-new')

    # Get the HTML content of the table
    table_html = table_element.get_attribute('outerHTML')

    # Find the next button element
    next_button = driver.find_element(By.CSS_SELECTOR, 'a.screener-pages.is-next')

    # Save the HTML content of the first page as a text file
    with open(f'table_data_{tabele_no}.txt', 'w') as file:
        file.write(table_html)

    # Loop through the pages a number of times based on the total length
    for _ in range(total_data_tables_legth - 1):
        # Click on the next button
        next_button.click()
        time.sleep(5)  # Wait for the page to load

        tabele_no += 1

        # Find the table element
        table_element = driver.find_element(By.CLASS_NAME, 'styled-table-new')

        # Get the HTML content of the current page
        table_html = table_element.get_attribute('outerHTML')

        # Save the HTML content of the current page as a text file
        with open(f'table_data_{tabele_no}.txt', 'w') as file:
            file.write(table_html)

        # Check if there is a next button
        next_button = driver.find_element(By.CSS_SELECTOR, 'a.screener-pages.is-next') if driver.find_elements(By.CSS_SELECTOR, 'a.screener-pages.is-next') else None


def parse_and_extract_data_from_txt_files():
    global total_data_tables_legth

    # Initialize an empty list to store individual DataFrames
    dfs = []

    print("Extracting Data from txt files and Saving to Excel File ......")
    # Iterate through all the text files
    for tabele_no in range(1, total_data_tables_legth + 1):
        # Read the contents of the data table file
        file_path = f'table_data_{tabele_no}.txt'
        if os.path.exists(file_path):
            with open(file_path, 'r') as file:
                content = file.read()

            # Parse the HTML content
            soup = BeautifulSoup(content, 'html.parser')

            # Find the table rows
            rows = soup.find_all('tr')

            # Extract the data from the table
            data = []
            for row in rows:
                columns = row.find_all('td')
                if columns:
                    ticker = columns[0].text.strip()
                    company = columns[1].text.strip()
                    high_52w = columns[2].text.strip()
                    volume = columns[3].text.strip()
                    price = columns[4].text.strip()
                    change = columns[5].text.strip()
                    data.append({
                        'Ticker': ticker,
                        'Company': company,
                        '52W High': high_52w,
                        'Volume': volume,
                        'Price': price,
                        'Change': change
                    })

            # Convert the data into a DataFrame
            df = pd.DataFrame(data)

            # Append the DataFrame to the list
            dfs.append(df)

            # Remove the text file
            os.remove(file_path)

    # Concatenate all DataFrames in the list into a single DataFrame
    combined_df = pd.concat(dfs, ignore_index=True)

    # Write the combined DataFrame to an Excel file
    combined_df.to_excel('output.xlsx', index=False)

    print("Data Saved to Excel File Successfully ......")


def main():
    apply_filters_and_conditions()
    find_total_data_tables()
    fetch_and_save_data_tables_html_content()
    parse_and_extract_data_from_txt_files()


if __name__ == '__main__':
    main()
