import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from tqdm import tqdm
import re

chrome_options = Options()
chrome_options.add_argument("--window-size=1920x1080")
chrome_options.add_argument("--incognito")

driver = webdriver.Chrome(options=chrome_options,
                          executable_path='H:/DISTR/ChromeDriver/chromedriver_win32/chromedriver.exe')

main_url = 'https://apps.webofknowledge.com/Search.do?product=WOS&SID=D2ViYOZnYEkWrdvdzUQ&search_mode=GeneralSearch&prID=73c524d5-3513-47bb-bfc2-49ca653c100b'
driver.get(main_url)

WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, '.citation-report-summary-link')))
driver.find_element_by_css_selector('.userCabinet .nav-item~ .nav-item+ .nav-item .nav-link').click()
driver.find_element_by_css_selector('.en_US').click()
WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, '.citation-report-summary-link')))
link_to_cit = driver.find_element_by_css_selector('.citation-report-summary-link').get_attribute('href')
print(link_to_cit)
driver.get(link_to_cit)

full_data = pd.DataFrame()

link_parts = re.split(r'&', link_to_cit)
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#pageCount\.bottom')))
page_count = int(driver.find_element_by_css_selector('#pageCount\.bottom').text)
for k in tqdm(range(1, page_count + 1, 1)):
    # Needed page getting
    for j in range(len(link_parts)):
        if 'page' in link_parts[j]:
            link_parts[j] = 'page={}'.format(k)
    next_page = '&'.join(link_parts)
    driver.get(next_page)
    # Extracting all cits on page
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.CitReportTotalRow1 .tcPerYear')))
    years = driver.find_elements_by_css_selector('.CitReportTotalRow1 .tcPerYear')
    years = [j.text for j in years]
    if years[0] != '2014':
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.search-results-item .tcPerYear')))
        cits_2017_2021 = driver.find_elements_by_css_selector('.search-results-item .tcPerYear')
        cits_2017_2021 = pd.Series([j.text for j in cits_2017_2021])
        driver.find_element_by_css_selector('.snowplow-cited-rep-citation-report-previous-set-of-years img').click()
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.search-results-item .tcPerYear')))
        cits_2014_2016 = driver.find_elements_by_css_selector('.search-results-item .tcPerYear')
        cits_2014_2016 = pd.Series([j.text for j in cits_2014_2016])
    else:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.search-results-item .tcPerYear')))
        cits_2014_2016 = driver.find_elements_by_css_selector('.search-results-item .tcPerYear')
        cits_2014_2016 = pd.Series([j.text for j in cits_2014_2016])
        driver.find_element_by_css_selector('a img').click()
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.search-results-item .tcPerYear')))
        cits_2017_2021 = driver.find_elements_by_css_selector('.search-results-item .tcPerYear')
        cits_2017_2021 = pd.Series([j.text for j in cits_2017_2021])
    # Extracting urls
    WebDriverWait(driver, 20).until(EC.presence_of_element_located(
        (By.XPATH, '//a[@class="smallV110 snowplow-cited-rep-citation-report-articlename"]')))
    doc_urls = driver.find_elements_by_xpath(
        '//a[@class="smallV110 snowplow-cited-rep-citation-report-articlename"]')
    doc_urls = pd.Series([j.get_attribute('href') for j in doc_urls])
    # Making DF from cit lists and doc_urls
    cits_2017_dict = {}
    cits_2014_dict = {}
    for j in range(len(doc_urls)):
        url = str(doc_urls[j])
        cit_values = list(cits_2017_2021.loc[j * 5:(j * 5 + 4)])
        cits_2017_dict[url] = cit_values
        cit_values = list(cits_2014_2016.loc[j * 3:(j * 3 + 2)])
        cits_2014_dict[url] = cit_values

    cits_2017_df = pd.DataFrame(data=cits_2017_dict)
    cits_2017_df = cits_2017_df.T.reset_index(drop=False)
    print(cits_2017_df)
    cits_2017_df.columns = ['link', 'y_2017', 'y_2018', 'y_2019', 'y_2020', 'y_2021']

    cits_2014_df = pd.DataFrame(data=cits_2014_dict)
    cits_2014_df = cits_2014_df.T.reset_index(drop=False)
    cits_2014_df.columns = ['link', 'y_2014', 'y_2015', 'y_2016']

    cits_full_page = cits_2014_df.join(cits_2017_df.set_index('link'), on='link')
    print(cits_full_page)
    page_docs_data = pd.DataFrame()
    for i in tqdm(range(len(doc_urls))):
        driver.get(doc_urls[i])
        try:
            driver.find_element_by_css_selector('#show_more_authors_authors_txt_label a').click()
            more_clicked_flag = 1
        except:
            more_clicked_flag = 0
            pass
        if more_clicked_flag == 0:
            authors = driver.find_elements_by_css_selector('.title+ .block-record-info .FR_field a')
            authors = pd.Series([el for el in authors if (not el.text.isdigit())
                                 and (el.text != '') and (el.text != '...Меньше')])
            links_unique = pd.Series([author_link.get_attribute('href') for author_link in authors])
            authors = driver.find_elements_by_css_selector('.title+ .block-record-info .FR_field a')
            authors = pd.Series([el.text for el in authors])
        else:
            authors = driver.find_elements_by_css_selector('.title+ .block-record-info .FR_field a')
            authors = pd.Series([el for el in authors if (not el.text.isdigit())
                                 and (el.text != '') and (el.text != '...Less')])
            links_unique = pd.Series([author_link.get_attribute('href') for author_link in authors])
            authors = driver.find_elements_by_css_selector('.title+ .block-record-info .FR_field a')
            authors = [el.text for el in authors]
            for l_ind in range(-1, len(authors) - 1, 1):
                if authors[l_ind] == '' or authors[l_ind] == '...Less':
                    del authors[l_ind]
            authors = pd.Series(authors)

        names = []
        au_aff_numbers = []
        links = []
        local_numbers = []
        for j in range(len(authors)):
            if j != len(authors) - 1:
                if (not authors[j].isdigit()) and (not authors[j + 1].isdigit()):
                    names.append(authors[j])
                    au_aff_numbers.append('-1')
                if (not authors[j].isdigit()) and authors[j + 1].isdigit():
                    for m in range(j + 1, len(authors), 1):
                        if authors[m].isdigit():
                            local_numbers.append(authors[m])
                        else:
                            break
                    for num in local_numbers:
                        names.append(authors[j])
                        au_aff_numbers.append(num)
                local_numbers = []
            else:
                if not authors[j].isdigit():
                    names.append(authors[j])
                    au_aff_numbers.append('-1')

        link_counter = 0
        print(authors)
        print(links_unique)
        for j in range(len(authors)):
            if j != len(authors) - 1:
                if (not authors[j].isdigit()) and authors[j + 1].isdigit():
                    for m in range(j + 1, len(authors), 1):
                        if authors[m].isdigit():
                            links.append(links_unique[link_counter])
                        else:
                            link_counter += 1
                            break
                if (not authors[j].isdigit()) and (not authors[j + 1].isdigit()):
                    links.append(links_unique[link_counter])
                    link_counter += 1
            else:
                if not authors[j].isdigit():
                    links.append(links_unique[link_counter])

        try:
            aff_add = driver.find_elements_by_css_selector('.FR_table_noborders:nth-child(9) .fr_address_row2')
            aff_add = [uni.text for uni in aff_add]
            aff_number = []
            aff_uni = []
            for j in range(len(aff_add)):
                aff_number.append(re.findall(r'\d+', aff_add[j])[0])
                aff_uni.append(re.split(r'] ', aff_add[j])[1])
                if 'Satbayev' in aff_uni[j]:
                    aff_uni[j] = 'Satbayev University'
        except:
            aff_add = driver.find_elements_by_css_selector('.FR_table_noborders:nth-child(12) .fr_address_row2')
            aff_add = [uni.text for uni in aff_add]
            aff_number = []
            aff_uni = []
            for j in range(len(aff_add)):
                aff_number.append(re.findall(r'\d+', aff_add[j])[0])
                aff_uni.append(re.split(r'] ', aff_add[j])[1])
                if 'Satbayev' in aff_uni[j]:
                    aff_uni[j] = 'Satbayev University'
        if aff_number == []:
            try:
                aff_add = driver.find_elements_by_css_selector('.FR_table_noborders:nth-child(15) .fr_address_row2')
                aff_add = [uni.text for uni in aff_add]
                aff_number = []
                aff_uni = []
                for j in range(len(aff_add)):
                    aff_number.append(re.findall(r'\d+', aff_add[j])[0])
                    aff_uni.append(re.split(r'] ', aff_add[j])[1])
                    if 'Satbayev' in aff_uni[j]:
                        aff_uni[j] = 'Satbayev University'
            except:
                aff_add = driver.find_elements_by_css_selector('.FR_table_noborders:nth-child(18) .fr_address_row2')
                aff_add = [uni.text for uni in aff_add]
                aff_number = []
                aff_uni = []
                for j in range(len(aff_add)):
                    aff_number.append(re.findall(r'\d+', aff_add[j])[0])
                    aff_uni.append(re.split(r'] ', aff_add[j])[1])
                    if 'Satbayev' in aff_uni[j]:
                        aff_uni[j] = 'Satbayev University'
        aff_table = pd.DataFrame(data={'aff_number': aff_number, 'aff_uni': aff_uni})
        authors_table = pd.DataFrame(data={'author': names, 'author_link': links, 'aff_number': au_aff_numbers})
        au_full = authors_table.join(aff_table.set_index('aff_number'), on='aff_number').drop(['aff_number'], axis=1)

        # Subject adding
        doc_ad_info = driver.find_elements_by_css_selector('#journal_section+ .block-record-info .FR_field')
        doc_ad_info = pd.Series([j.text for j in doc_ad_info])
        print(doc_ad_info)
        subjects_flag = 0
        for j in range(len(doc_ad_info)):
            if 'Categories' in doc_ad_info[j]:
                au_full['subject'] = re.split(r'Categories:', doc_ad_info[j])[1]
                subjects_flag = 1
        if subjects_flag == 0:
            au_full['subject'] = ''

        # Doc_wos_id adding
        try:
            WebDriverWait(driver, 1).until(
                EC.text_to_be_present_in_element((By.CSS_SELECTOR, '#hidden_section_label'),
                                                 'See more data fields'))
            driver.find_element_by_css_selector('#hidden_section_label').click()
        except:
            pass
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#hidden_section .FR_field+ .FR_field value')))
        doc_ad_info = driver.find_elements_by_css_selector('#hidden_section .FR_field+ .FR_field value')
        doc_ad_info = [j.text for j in doc_ad_info if 'WOS' in j.text][0]
        au_full['doc_wos_id'] = doc_ad_info

        # Title adding
        title = driver.find_element_by_css_selector('.title value').text
        au_full['title'] = title

        # Link adding
        au_full['link'] = doc_urls[i]

        # Source title adding
        source_title = driver.find_element_by_css_selector('.sourceTitle value').text
        au_full['source_title'] = source_title

        # DOI, publ_date, doc_type adding
        block_record_info = driver.find_elements_by_css_selector('.block-record-info-source .FR_field')
        block_record_info = pd.Series([j.text for j in block_record_info])
        print(block_record_info)

        doi_flag = 0
        date_flag = 0
        type_flag = 0
        for j in range(len(block_record_info)):
            try:
                if 'DOI' in block_record_info[j]:
                    au_full['doi'] = re.split(r'DOI: ', block_record_info[j])[-1]
                    doi_flag = 1
            except:
                doi_flag = 0
            try:
                if 'Published' in block_record_info[j]:
                    date = re.split(r' ', re.split(r'Published: ', block_record_info[j])[1])
                    au_full['publ_date'] = date[-1]
                    date_flag = 1
            except:
                date_flag = 0
            try:
                if 'Type' in block_record_info[j]:
                    au_full['doc_type'] = re.split(r'Document Type:', block_record_info[j])[-1]
                    type_flag = 1
            except:
                type_flag = 0

        if doi_flag == 0:
            au_full['doi'] = ''
        if date_flag == 0:
            au_full['publ_date'] = ''
        if type_flag == 0:
            au_full['doc_type'] = ''

        # ISSN, eISSN adding
        block_record_info = driver.find_elements_by_css_selector('.block-record-info:nth-child(1) .FR_field')
        block_record_info = pd.Series([j.text for j in block_record_info])
        print(block_record_info)

        issn_flag = 0
        eissn_flag = 0
        for j in range(len(block_record_info)):
            try:
                if 'ISSN' in block_record_info[j]:
                    au_full['ISSN'] = re.split(r'ISSN: ', block_record_info[j])[-1]
                    issn_flag = 1
            except:
                issn_flag = 0
            try:
                if 'eISSN' in block_record_info[j]:
                    au_full['eISSN'] = re.split(r'eISSN: ', block_record_info[j])[-1]
                    eissn_flag = 1
            except:
                eissn_flag = 0
        if issn_flag == 0:
            au_full['ISSN'] = ''
        if eissn_flag == 0:
            au_full['eISSN'] = ''

        # Searching for author's WOS ID
        au_wos_ids = []
        for j in range(len(links_unique)):
            driver.get(links_unique[j])
            try:
                WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.wat-author-record__rid')))
                au_wos_ids.append(driver.find_element_by_css_selector('.wat-author-record__rid').text)
            except:
                au_wos_ids.append('')
        au_wos_ids = pd.Series(au_wos_ids)
        wos_id_table = pd.DataFrame(data={'author_link': links_unique, 'au_wos_id': au_wos_ids})
        au_full = au_full.join(wos_id_table.set_index('author_link'), on='author_link')

        # Storaging all the doc's data
        print(au_full)
        page_docs_data = page_docs_data.append(au_full)

    # Joining data of one page
    page_full_data = page_docs_data.join(cits_full_page.set_index('link'), on='link')
    page_full_data = page_full_data.reset_index(drop=True)
    print(page_full_data)

    # Separating subjects of page
    page_full_data_sep = {}
    for j in range(len(page_full_data)):
        if not page_full_data['subject'].isna()[j] and page_full_data.loc[j, 'subject'] != '':
            subj_sep = pd.Series(re.split(r';', page_full_data.loc[j, 'subject']))
            subj_sep = subj_sep.dropna().reset_index(drop=True)
        else:
            subj_sep = pd.Series(data='')
        column_series = pd.Series(page_full_data.columns)
        for m in range(len(subj_sep)):
            for n in range(len(column_series)):
                if column_series[n] == 'subject':
                    if column_series[n] in page_full_data_sep.keys():
                        page_full_data_sep[column_series[n]].append(subj_sep[m])
                    else:
                        page_full_data_sep[column_series[n]] = [subj_sep[m]]
                else:
                    if column_series[n] in page_full_data_sep.keys():
                        page_full_data_sep[column_series[n]].append(page_full_data.loc[j, column_series[n]])
                    else:
                        page_full_data_sep[column_series[n]] = [page_full_data.loc[j, column_series[n]]]
    page_full_data_sep = pd.DataFrame(page_full_data_sep)
    print(page_full_data_sep)

    # full_data appending
    full_data = full_data.append(page_full_data_sep)
    print(full_data)
    if k % 10 == 0:
        full_data.to_excel('wos_full_data_130421.xlsx', index=False)

full_data.to_excel('wos_full_data_130421.xlsx', index=False)
driver.quit()