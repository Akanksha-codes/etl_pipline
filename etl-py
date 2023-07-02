from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import Select
import pandas as pd
from selenium.common.exceptions import NoSuchElementException
import csv
import pytesseract
from PIL import Image


fault_pages=[]
fault_rows=[]
driver = webdriver.Chrome()
driver.get("https://www.bundesanzeiger.de/pub/de/start")
time.sleep(2)
# Cookies accept
btn_acc_cookies = driver.find_element(By.ID, "cc_all")
btn_acc_cookies.click()
time.sleep(2)

# Bundesministerium search
input_search_keyword = driver.find_element(By.ID, "id3")
input_search_keyword.send_keys("a")


#Uncomment following
#input_search_keyword.send_keys(".,.,")


dropdown = driver.find_element(By.CLASS_NAME, "select2-selection__arrow")

# Click on the dropdown to open it
dropdown.click()
time.sleep(1)

# # Wait for the "Amtlicher Teil" option to be visible
option = driver.find_element(By.XPATH, '//li[contains(text(), "Alle Bereiche")]')

# Click on the "Amtlicher Teil" option
option.click()
time.sleep(1)

# Press the enter button
input_search_keyword.submit()

# On the new page
# Select the 100 rows per page option button
select_element = driver.find_element(By.NAME, "hitsperpage-select")
select = Select(select_element)
select.select_by_value("argus-HitsPerPage100")
time.sleep(2)

# Find the div element with class "page_count"
num_of_pages_element = driver.find_element(By.CLASS_NAME, "page_count")

# Find the span element within the div
span_element = num_of_pages_element.find_element(By.TAG_NAME, "span")

# Extract the text content of the span element
total_pages = int(span_element.text)
print("Total Pages:", total_pages)

# Create a Pandas Excel writer
excel_writer = pd.ExcelWriter('alldata.xlsx', engine='xlsxwriter')
min_page_num=1
max_pages=1
try:
    for this_page in range(1, max_pages+1):
        # Mechanism to click on the current page
        title_string = "//a[@title='Zur Seite " + str(this_page) + "']"
        page_num_btn = driver.find_element(By.XPATH, title_string)
        page_num_btn.click()
        time.sleep(0.2)
        if this_page>=min_page_num:
            myhtml = driver.page_source
            soup = BeautifulSoup(myhtml, 'html.parser')
            row_back_div = soup.find_all(class_=["row", "row back"])

            names_div_shortlist = []
            for f in row_back_div:
                find_first = f.find(class_="col-md-3")
                names_div_shortlist.append(find_first)
            names = []
            for r in names_div_shortlist:
                if r is not None:
                    name_div = r.find(class_="first")
                    if name_div is not None:
                        value = name_div.get_text(strip=True)
                        names.append(value)

            part_div_shortlist = []
            for f in row_back_div:
                find_part = f.find(class_="col-md-2")
                part_div_shortlist.append(find_part)
            parts = []
            for r in part_div_shortlist:
                if r is not None:
                    part_div = r.find(class_="part")
                    if part_div is not None:
                        value = part_div.get_text(strip=True)
                        parts.append(value)

            info_div_shortlist = []
            links = []
            for f in row_back_div:
                find_info = f.find(class_="col-md-5")
                info_div_shortlist.append(find_info)
            infos = []
            for r in info_div_shortlist:
                if r is not None:
                    info_div = r.find(class_="info")
                    if info_div is not None:
                        value = info_div.get_text(strip=True)
                        infos.append(value)

            date_div_shortlist = []
            for f in row_back_div:
                find_date = f.find_all(class_="col-md-2")
                if len(find_date) >= 2:
                    date_div_shortlist.append(find_date[1])
            dates = []
            for r in date_div_shortlist:
                if r is not None:
                    date_div = r.find(class_="date")
                    if date_div is not None:
                        value = date_div.get_text(strip=True)
                        dates.append(value)

            div_htmls = []
            # Find all div elements with the class "info" to click on them and get data on those pages
            div_elements = driver.find_elements(By.XPATH, '//div[contains(@class, "info")]')
            pdf_download_links = []
            pdf_contents = []
            print(len(div_elements))
            # Iterate over the div elements
            for x in range(16, len(div_elements)):

                #To extract collapsable links
                for i in range(1, len(div_elements)-100-1):
                    try:
                        link_element2 = div_elements[x-i].find_element(By.CLASS_NAME, "toggle-link.collapsed.argus-A41")
                        link_element2.click()
                        div_elements = driver.find_elements(By.XPATH, '//div[contains(@class, "info")]')
                        
                    except NoSuchElementException:
                        pass 


                try:
                    div_elements = driver.find_elements(By.XPATH, '//div[contains(@class, "info")]')
                    
                    link_element = div_elements[x].find_element(By.XPATH, './/a')
                except NoSuchElementException:
                    pass 

                
                
                if link_element is not None:
                    link_element.click()
                    time.sleep(0.2)
                    # try:

                    #     # Capture a screenshot of the captcha wrapper element
                    #     captcha_wrapper = driver.find_element(By.CLASS_NAME,"captcha_wrapper")
                    #     captcha_wrapper.screenshot("captcha.png")
                    #     print(captcha_wrapper.get_attribute("innerHTML"))

                    #     # Use pytesseract to extract text from the captcha image
                    #     captcha_image = Image.open("captcha.png")
                    #     captcha_image.save("captcha.png")
                    #     captcha_text = pytesseract.image_to_string(captcha_image)
                    #     print("captcha_text")
                    #     print(captcha_text)
                    #     print("captcha_text 123")

                    #    # time.sleep(20)
                    #     # Fill the captcha solution into the input field
                    #     solution_input = driver.find_element(By.NAME,"solution")
                    #     solution_input.send_keys(captcha_text)

                    #     # Click the OK button
                    #     ok_button = driver.find_element(By.NAME,"confirm-button")
                    #     ok_button.click()
                    # except NoSuchElementException:
                    #     pass 



                    # Find the <a> element with the class "fa-file-download"
                # download_button = driver.find_element(By.LINK_TEXT, "Publikation als PDF herunterladen")
                   # print("captcha_text dddddd")
                    
                    download_button = None

                    pdf_content= None
                    try:
                        download_button = driver.find_element(By.LINK_TEXT, "Publikation als PDF herunterladen")
                    except NoSuchElementException:
                        fault_pages.append(this_page)
                        fault_rows.append(x)
                        pass
                    try:
                        pdf_content = driver.find_element(By.CLASS_NAME, "publication_container")
                    except NoSuchElementException:
                        pass
                    if pdf_content:
                            innerhtml = pdf_content.get_attribute("innerHTML")
                            soup = BeautifulSoup(innerhtml, 'html.parser')
                            text = soup.get_text(separator=' ')
                            pdf_contents.append(text.strip())
                    else:
                        pdf_contents.append("Maybe captcha problem")

                    if download_button:
                        # Get the href attribute value
                        download_link = download_button.get_attribute("href")
                        if download_link:
                            pdf_download_links.append(download_link)
                        else:
                            pdf_download_links.append("")
                    else:
                        # fault_pages.append(this_page)
                        # fault_rows.append(x)
                       
                        pdf_download_links.append("")
                    driver.back()
                    time.sleep(0.2)
                    # Re-locate the div elements again after going back
                    div_elements = driver.find_elements(By.XPATH, '//div[contains(@class, "info")]')
                    # if x == 1:
                    #     break

            # Create a data frame for the current page's data
            print(len(names))
            print(len(parts))
            print(len(infos))
            print(len(dates))
            print(len(pdf_contents))
            print(len(pdf_download_links))
            data = {
                'Name': names,
                'Part': parts,
                'Info': infos,
                'Date': dates,
                'PDF Contents': pdf_contents,
                'PDF Download Link': pdf_download_links
            }
            df = pd.DataFrame(data)

            # Write the data frame to a new sheet in the Excel file
            sheet_name = 'Page ' + str(this_page)
            df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

            print("Data for Page", this_page, "saved to CSV")

except Exception as e:
    print("An error occurred:", str(e))

finally:
    # Save and close the Excel writer
    excel_writer.save()
    excel_writer.close()

    # Close the browser
    driver.quit()

    # Save fault data to CSV
    fault_data = list(zip(fault_pages, fault_rows))
    fault_csv_file = 'fault_data.csv'
    with open(fault_csv_file, 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['fault_pages', 'fault_rows'])
        writer.writerows(fault_data)
