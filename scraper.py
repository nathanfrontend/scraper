import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import os
import openpyxl


# Get today's date
today = datetime.today()
# Format the date as "dd/mm/yyyy"
formatted_date = today.strftime('%d/%m/%Y')
url_header = 'https://www.thegazette.co.uk/insolvency/notice?text=&categorycode=G405010102&categorycode=G405010202&categorycode=G405010302&categorycode=G405010501&categorycode=G405010502&categorycode=-2&location-postcode-1=&location-distance-1=1&location-local-authority-1=&numberOfLocationSearches=1&start-publish-date=23/01/2024&end-publish-date=&edition=&london-issue=&edinburgh-issue=&belfast-issue=&sort-by=&results-page-size=100&results-page='   

headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

def scrapeInitialLink(url):
        response = requests.get(url, headers=headers)
        # response.html.render(sleep=1)
        s = BeautifulSoup(response.text, 'html.parser')

        return s
       
def scrapeContent(links):

    

    
    data = []       
    for  url in links: 
        current_object = {}
        glob_url = url['link_gazzette']
        response = requests.get(glob_url, headers=headers) 
        s = BeautifulSoup(response.text, 'html.parser') 
        content = s.find('article')
        p =  content.find_all('p')
        noticeID = s.find('dd', {'property': 'gaz:hasNoticeID'}).text
        companyName = s.find('span', {'property': 'gazorg:name'}).text
        companyNumber = s.find('span', {'property': 'gazorg:companyNumber'}).text
        current_object['Gazzete'] = f'https://www.thegazette.co.uk/notice/{noticeID}' 
        current_object['Name of Company'] = companyName 
        current_object['Company Number'] = companyNumber

        for para in p:
               
            parts = para.text.split(':', 1)
            if len(parts) > 1:
                key = parts[0]
                value = parts[1].replace(':', '')
          
                current_object[key.strip()] = value.strip()
                
        data.append(current_object)
 
        

        for item in data:          
            if 'Company Number' in item:
                companyNo = item['Company Number']
                item['Company Site'] = f'https://www.thegazette.co.uk/company/{companyNo}'
                item['Companies House'] = f'https://find-and-update.company-information.service.gov.uk/company/{companyNo}'
            # else:
            #     item['Name of Company'] = 
            
       
                
                
                    

    # print(data)
    notices = pd.DataFrame(data)
   
    with pd.ExcelWriter('notices.xlsx', engine='openpyxl') as writer:
        pd.DataFrame(notices).to_excel(writer, sheet_name='Companies Yesterday', index=False)
        

       
        
def scrapeAllPages():
    url_data = main_func()
    data = []
    for url in url_data:
        # print(url)
        response = requests.get(url, headers=headers)
        # response.render(sleep=1)
        s = BeautifulSoup(response.text, 'html.parser')
        results = s.find(id='search-results')
        content = results.find_all('div', class_='feed-item')
        
      
        for pub_feed in content:
            p = pub_feed.find('p')
            a = pub_feed.find('a')
            data.append({'link_gazzette':'https://www.thegazette.co.uk' + a.get('href')})
        
    scrapeContent(data)

       
   
def getnextpage(soup):
    # this will return the next page URL  
    pages = soup.find('div', {'class': 'nav-container feed-pagination'})

    if pages:
        next_page_link = pages.find('li', {'class':'next'}).find('a')
        if next_page_link and 'href' in next_page_link.attrs:
            url = next_page_link['href']
            return url
        else:
            url = 'no link'
            return url
    else:
        return
        
def main_func():
    url = f'{url_header}0'       
    url_data = []
    while True:
        data = scrapeInitialLink(url)
        url = getnextpage(data)
        if not url:
            url = f'{url_header}1'
            url_data.append(url)  
            break
        if url == 'no link':      
            break
        url_data.append(url)     
    return url_data   

scrapeAllPages()


      

               
 
