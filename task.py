"""Template robot with Python."""

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Tables
import os
import time


class Challenge_Class():
    
    def __init__(self):
        
        self.browser_lib = Selenium()
    
        self.excel_lib = Files()
        
        self.tables = Tables()
        
        self.agencies_name = []
        
        self.agencies_spending = []
        
        self.webelements = []
        
        self.links = []
        
        self.urls = []
        
        

    def open_the_website(self, url):
        
        self.browser_lib.set_download_directory(os.getcwd()+'/output/')
        
        self.browser_lib.open_available_browser(url)
        
        self.browser_lib.maximize_browser_window()
    
    
    def extract_Total_Spendings(self):
        
        self.excel_lib.create_workbook("output/Amounts.xlsx",fmt='xlsx')
        
        self.excel_lib.rename_worksheet(self.excel_lib.get_active_worksheet(), "Agencies")
        
        self.excel_lib.set_cell_value(1, 1, 'Agencies')
        
        self.excel_lib.set_cell_value(1, 2, 'Total FY2021 Spending:')
        
        self.browser_lib.click_link("#home-dive-in", False)
        
        self.browser_lib.wait_until_page_contains("Total")
        
        self.webelements = self.browser_lib.get_webelements("xpath: //*[text()[contains(.,'Total')]]")
        
        for i, val in enumerate(self.webelements):
            
            self.agencies_name.append(self.browser_lib.get_text(self.webelements[i]).splitlines()[0])
            
            self.agencies_spending.append(self.browser_lib.get_text(self.webelements[i]).splitlines()[2])
            
            self.excel_lib.set_cell_value(i+2, 1, self.agencies_name[i])
            
            self.excel_lib.set_cell_value(i+2, 2, self.agencies_spending[i]) 
            
        self.excel_lib.save_workbook("output/Amounts.xlsx") 
        
    
    def extract_Individual_Investments(self):
        
        self.excel_lib.create_worksheet('Individual Investments')
        
        self.browser_lib.click_element(self.webelements[24])
        
        self.browser_lib.wait_until_page_contains_element("xpath: //*[contains(text(),'Investment Title')]", 10)
        
        self.browser_lib.select_from_list_by_value(self.browser_lib.get_webelement('name:investments-table-object_length'), "-1")
        
        time.sleep(15)
        
        table = self.browser_lib.execute_javascript("const data = [[],[],[],[],[],[],[]];\
                                                    var table = document.getElementById('investments-table-object'); \
                                                    for (var i = 0, row; row = table.rows[i]; i++) \
                                                    {for (var j = 0, col; col = row.cells[j]; j++) \
                                                    {data[j][i] = col.innerText } }; \
                                                    return data")
        
        new_table = list(zip(*table)); new_table.pop(0)
        
        self.excel_lib.append_rows_to_worksheet(new_table, 'Individual Investments')
        
        self.excel_lib.save_workbook("output/Amounts.xlsx") 
        
        self.links = table[0][2:]
        
        
        
    def download_pdfs(self):
        
        for link in self.links:
            
            href = self.browser_lib.get_element_attribute("xpath: //*[contains(text(),'" + link + "')]", 'href')
            
            if href:   
                
                self.urls.append(href)         
                
        for url in self.urls:
            
            try:
            
                self.browser_lib.go_to(url)
            
                self.browser_lib.wait_until_page_contains_element("xpath: //*[contains(text(),'Download Business Case PDF')]", 10)
            
                self.browser_lib.click_link("xpath: //*[contains(text(),'Download Business Case PDF')]")
                
                self.browser_lib.wait_until_page_does_not_contain("xpath: //*[contains(text(),'Generating')]")
            
                time.sleep(10)
                
            except: pass
        
        
    
def main():
    
    try:
        
        challenge = Challenge_Class()
        
        challenge.open_the_website("http://itdashboard.gov/")
        
        challenge.extract_Total_Spendings()
        
        challenge.extract_Individual_Investments()
        
        challenge.download_pdfs() 
        
        
    finally:
        
        challenge.browser_lib.close_all_browsers()




if __name__ == "__main__":
    main()