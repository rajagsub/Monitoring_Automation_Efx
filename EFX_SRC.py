"""
    EFX.py, by Rajagopalan S(AF40411), , 2021-04-20.

    This program make a request to the EFX portal with the given list of PRODUCER-ID's
    and gets the Failed/Error transfers and saves it into the excel file.

    Pre-requisite : Need to import below libraries with python ver 3.8.8

    Input : Uses the Login.txt file with the list of the Producer-ID & Password.

    Output: EFX_Routing_and_Failed_Files_Report.xlsx which has the list of files that are in FAILED/ERROR state and need
    attention.

    #raj1008 Date: 10/08/2021 - Fixed the total Completed processed counts by subtracting tot_rec - Err_rec.
"""

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from datetime import date
import time
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd


def get_data(self, tot_count_a):
    """
        Parameter:  Total count of the records present in the EFX.

        Function:   Loops thru until the scroll reaches the total count of the records present in EFX portal
                    and fetch total records from the EFX portal, Filtered data other than Routed and Complete.

        Return:     List of consolidated records (full_data_x), List of Filtered Data (filtered_data_x),
                    List of Routed records (full_complete_data_x).

    """

    # initialize the list & variables used to fetch.
    raw_data = []
    a = 0
    init = 1
    adder = 51
    while True:
        if init > tot_count_a:
            # self.driver.get_screenshot_as_file("screenshot.png")
            break
        else:
            # Wait a page to load
            time.sleep(2)
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'lxml')
            div_elements = soup.find('div', id='isc_9N')
            tbody = div_elements.find('tbody')
            for i in range(init, adder):
                if init <= tot_count_a:
                    tr = tbody.find('tr', {"aria-posinset": init})
                    if tr == None:
                        print("Break at this:" , i)
                        break
                    else:
                        field = tr.get_text(separator=",")
                        raw_data.append(field)
                        init += 1
            adder += 50
            a += 1000
            # Scroll a page with 500 pxl
            # self.driver.get_screenshot_as_file("screenshotx.png")
            java_script = "document.querySelector('#isc_9N').scrollTop=" + str(a)
            self.driver.execute_script(java_script)
            # self.driver.get_screenshot_as_file("screenshoty.png")

    full_data = [str(i) for i in raw_data]

    tot_rec_cnt = len(full_data)
    tot_procced_cnt = 0

    full_data_x = []
    filtered_data_x = []
    full_complete_data_x = []

    for i in full_data:
        temp_list = i.split(',')
        full_data_x.append(temp_list[0])
        full_data_x.append(temp_list[1])
        full_data_x.append(temp_list[2])
        full_data_x.append(temp_list[3])
        full_data_x.append(temp_list[4])

        # if temp_list[0] == "Routed" and temp_list[1] == "Complete":
        if temp_list[0] == "Routed":
            tot_procced_cnt += 1
            full_complete_data_x.append(temp_list[0])
            full_complete_data_x.append(temp_list[1])
            full_complete_data_x.append(temp_list[2])
            full_complete_data_x.append(temp_list[3])
            full_complete_data_x.append(temp_list[4])

        if temp_list[0] != "Routed" or temp_list[1] != "Complete":
            filtered_data_x.append(temp_list[0])
            filtered_data_x.append(temp_list[1])
            filtered_data_x.append(temp_list[2])
            filtered_data_x.append(temp_list[3])
            filtered_data_x.append(temp_list[4])

    return full_data_x, full_complete_data_x, filtered_data_x, tot_rec_cnt, tot_procced_cnt

def data_write(datax, filename):
    """
        Parameter: List of Filtered data which needs to be written to the excel file with the file name of the file

        Function:  Creates the data to the dataframes and write the values to the excel and save it

        Return:    None

    """
    df = pd.DataFrame()

    # Creating columns to write into excel
    df['Arrived Status'] = datax[0::5]
    df['Delivery Status'] = datax[1::5]
    df['Producer'] = datax[2::5]
    df['Original File Name'] = datax[3::5]
    df['Discovery Time'] = datax[4::5]

    # Converting to Excel
    df.to_excel(filename, index= False)

def updt_html_file(producer_id, tot_rec_c, tot_err_c, tot_procced_c, lines, i):
    """
    :param producer_id:
    :param tot_rec_c:
    :param tot_err_c:
    :param tot_procced_c:
    :param lines:
    :param i:
    :return: lines (returns the line with the updated HTML body to attach in EMAIL).

    """
    updt_line1 = '<p>Producer-Id : ' + producer_id +', Unprocessed EFX Count: ' + str(tot_err_c) +', Processed EFX count: '+str(tot_procced_c)+', Total Count: '+str(tot_rec_c)+'&nbsp;</p>'

    lines.insert(i, updt_line1)

    return lines

class MyBotHome:
    """
        Parameter:
                    __init__: URL, user_name, user_pass, start_dt
                    login() : None
                    scrape(): None
                    closebrowser() : None

        Function:
                    __init__: initialize the drivers and set the driver arguments and get the URL.
                    login() : Maps the user-id and pass with the id and pass the values to the destination fields.

                    scrape(): 1. Scrape the total number of records from the EFX portal
                              2. Scrape the web page with the scroll and get the data.

        Return:     scrape(): Returns the list of the filtered data. (filtered_data)

    """
    def __init__(self, url, user_name, user_pass, start_dt):
        self.url = url
        self.user_name = user_name
        self.user_pass = user_pass
        self.start_dt = start_dt
        self.options = webdriver.ChromeOptions()
        self.options.headless = True
        self.options.add_argument("--window-size=1920,1080")
        self.options.add_argument('--ignore-certificate-errors')
        self.driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), options=self.options)
        self.driver.implicitly_wait(30)
        self.driver.get(self.url)

    def login(self):
        user_name_elem = self.driver.find_element_by_id("isc_H")
        user_name_elem.clear()
        user_name_elem.send_keys(self.user_name)
        user_pass_elem = self.driver.find_element_by_id("isc_L")
        user_pass_elem.clear()
        user_pass_elem.send_keys(self.user_pass)
        user_pass_elem.send_keys(Keys.RETURN)
        self.driver.implicitly_wait(10)
        time.sleep(1)
        user_from_dt = self.driver.find_element_by_id('isc_58')
        user_from_dt.send_keys(self.start_dt)
        self.driver.implicitly_wait(10)
        user_to_dt = self.driver.find_element_by_id('isc_6F')
        user_to_dt.send_keys(self.start_dt)
        # user_to_dt.send_keys("04/25/2021")
        time.sleep(5)
        user_from_dt.send_keys(Keys.RETURN)
        time.sleep(5)

    def scrape(self):
        try:
            # Check for the scroll bar if the content is loaded.
            self.driver.find_element_by_id("isc_9Z")

            # Beautiful Soup call to get landing page and scrape the data
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'lxml')

            # Scraping the Total number records found
            div_tot_rec = soup.find('div', id='isc_8P')
            tot_tbody = div_tot_rec.find('tbody').tr.text
            tot_count_split = tot_tbody.split(':')
            tot_count_temp = tot_count_split[1]
            tot_count = int(tot_count_temp.strip())

            print('Total Count:', tot_count)

            full_data, full_complete_data, filtered_data, tot_rec_cnt, tot_procced_cnt = get_data(self, tot_count)

            return full_data, full_complete_data, filtered_data, tot_rec_cnt, tot_procced_cnt

        except NoSuchElementException as NSE:
            try:
                # Try for the single page data
                # Beautiful Soup call to get landing page and scrape the data
                page_source = self.driver.page_source
                soup = BeautifulSoup(page_source, 'lxml')

                # Scraping the Total number records found
                div_tot_rec = soup.find('div', id='isc_8P')
                tot_tbody = div_tot_rec.find('tbody').tr.text
                tot_count_split = tot_tbody.split(':')
                tot_count_temp = tot_count_split[1]
                tot_count = int(tot_count_temp.strip())

                print('Total Count:', tot_count)
                full_data, full_complete_data, filtered_data, tot_rec_cnt, tot_procced_cnt = get_data(self,tot_count)

                return full_data, full_complete_data, filtered_data, tot_rec_cnt, tot_procced_cnt

            except IndexError:
                print("No data for the selected date range")
                full_data_final = []
                full_complete_data = []
                filtered_data_final = []
                tot_rec_count = 0
                tot_procced_count = 0
                return full_data_final, full_complete_data, filtered_data_final, tot_rec_count, tot_procced_count

    def closebrowser(self):
        self.driver.quit()

def compare_failed_complete(failed_list, complete_list, tot_procced_count):
    """
    :param failed_list:
    :param complete_list:
    :param tot_procced_count:
    :return: failed_list_finally (List of final failed members), tot_procced_count(Updated Total Member Processed cnt).

    """

    before = 0
    after = 0
    inter_a = 0
    inter_b = 0
    failed_list_finally = []

    """ 
        1)  Loops through the FAILED LIST with the range of 3, len(failed_list), in the step of 5 to get 
            file name, E.G "QA_MO_AEC_20210410152157.TXT"
            
        2)  Loops through the COMPLETED LIST with the range of 3, len(failed_list), in the step of 5 to get 
            file name, E.G "QA_MO_AEC_20210410152157.TXT"
        
        3)  Check if Failed file present in Completed file list.
        
        4)  Check if the Failed file date is GREATER than Completed File date.
        
        5) If it mets above condition append it to the failed_list_finally LIST.
        
    """
    for x in range(3, len(failed_list) + 1, 5):
        for y in range(3, len(complete_list) + 1, 5):
            if failed_list[x] == complete_list[y]:
                # print("Match found,", failed_list[x])
                inter_a = x + 1
                inter_b = y + 1
                date_f_yyyy = failed_list[inter_a][6:10]
                date_c_yyyy = complete_list[inter_b][6:10]

                if (date_f_yyyy == date_c_yyyy):
                    if (failed_list[inter_a] < complete_list[inter_b]):
                        found_flag = True
                        # raj1008 tot_procced_count += 1
                        break
                elif (date_f_yyyy < date_c_yyyy):
                    found_flag = True
                    # raj1008 tot_procced_count += 1
                    break
                else:
                    found_flag = False
            else:
                found_flag = False

        if found_flag == False:
            before = x - 3
            after = x + 1
            for i in range(before, after + 1):
                failed_list_finally.append(failed_list[i])

    # raj1008return failed_list_finally, tot_procced_count
    return failed_list_finally

if __name__ == '__main__':

    current_day = date.today()
    start = time.time()
    print("Today's Date:", current_day)
    formatted_date = date.strftime(current_day, "%m/%d/%Y")

    all_producer_err = []
    all_producer_succ = []
    index = 3

    # Open Login.txt file and Email_body.html file.
    log_file = open("Login.txt", "r")
    temp_email_file = open('Email_Body.html', 'r')
    lines = temp_email_file.readlines()

    # Loops through the Login.txt file and process all the producers.
    for i in log_file:
        temp = i.split(",")
        user_id = temp[0]
        pass_id = temp[1].rstrip("\n")

        print(user_id)

        my_bot = MyBotHome('Masked URL', user_id, pass_id,
                           formatted_date)
                           #   '04/25/2021')

        my_bot.login()
        full_data_final, full_complete_data_final, filtered_data_final, tot_rec_count, tot_procced_count = my_bot.scrape()
        my_bot.closebrowser()

        # Check if the failed record found in the COMPLETED LIST
        # raj1008 failed_list_finally_x, tot_procced_counts = compare_failed_complete(filtered_data_final, full_complete_data_final, tot_procced_count)
        failed_list_finally_x = compare_failed_complete(filtered_data_final, full_complete_data_final, tot_procced_count)
        tot_err_count = int(len(failed_list_finally_x)/5)
        # raj1008
        tot_procced_counts = tot_rec_count - tot_err_count
        print('tot_procced_counts', tot_procced_counts)
        # end raj1008
        # Coping the Filtered data to another list for adding multiple producer
        if (failed_list_finally_x != None):
            all_producer_err = all_producer_err + failed_list_finally_x

        if (full_data_final!= None):
            all_producer_succ = all_producer_succ + full_data_final

        # Update HTML file for the EMAIL Body content.
        lines_val = updt_html_file(user_id, tot_rec_count, tot_err_count, tot_procced_counts,lines,index)

    # Close the Email_body.html file.
    temp_email_file.close()

    # Open Email_body1.html file to write EMAIL body content, and Close the file.
    file = open('Email_Body1.html', 'w')
    file.writelines(lines_val)
    file.close()

    err_file_name = 'EFX_Routing_and_Failed_Files_Report.xlsx'
    succ_file_name = 'EFX_Consolidated_Report.xlsx'

    # Write the Err & Succ list to the Excel file.
    data_write(all_producer_err, err_file_name)
    data_write(all_producer_succ, succ_file_name)

    log_file.close()
    end = time.time()
    print(f"Runtime of the program is {end - start}")
