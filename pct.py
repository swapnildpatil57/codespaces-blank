import requests
from bs4 import BeautifulSoup
import zipfile
from io import BytesIO
import openpyxl
from openpyxl import Workbook

# Function to scrape the data from a single page and save the HTML
def scrape_page(url, page_num, html_files):
    response = requests.get(url)
    html_content = response.text
    html_files.append((f"{page_num}.html", html_content))
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Extracting data from the table
    table = soup.find('table', class_='table')
    rows = table.find_all('tr', class_='team')
    
    page_data = []
    for row in rows:
        cols = row.find_all('td')
        data = {
            'year': (cols[1].text.strip()),
            'team': cols[0].text.strip(),
            'wins': int(cols[2].text.strip()),
            'losses': int(cols[3].text.strip()),
            'ot_losses': (cols[4].text.strip()),
            'win_percent': float(cols[5].text.strip()),
            'goals_for': int(cols[6].text.strip()),
            'goals_against': int(cols[7].text.strip()),
            'diff': int(cols[8].text.strip())
        }
        page_data.append(data)
    
    return page_data

# Function to scrape all pages and collect data
def scrape_all_pages(base_url):
    all_data = []
    html_files = []
    
    # Loop through all pages
    for page_num in range(1, 25):
        url = f"{base_url}?page_num={page_num}"
        page_data = scrape_page(url, page_num, html_files)
        all_data.extend(page_data)
    
    return all_data, html_files

# Function to save HTML files to a ZIP archive
def save_to_zip(html_files, zip_filename):
    with zipfile.ZipFile(zip_filename, 'w') as z:
        for filename, content in html_files:
            z.writestr(filename, content)

# Function to create Excel file with two sheets
def create_excel(all_data, excel_filename):
    wb = Workbook()
    
    # Sheet 1: NHL Stats 1990-2011
    sheet1 = wb.active
    sheet1.title = "NHL Stats 1990-2011"
    
    headers = ["Year", "Team", "Wins", "Losses", "OT Losses", "Win %", "Goals For", "Goals Against", "Goal Differential"]
    sheet1.append(headers)
    



    for data in all_data:
        if data['year'] == '1990' or data['year'] == '1991':
            sheet1.append([
                data['year'], data['team'], data['wins'], data['losses'], data['ot_losses'],
                data['win_percent'], data['goals_for'], data['goals_against'], data['diff']
            ])
    
    # Sheet 2: Winner and Loser per Year
    sheet2 = wb.create_sheet(title="Winner and Loser per Year")
    sheet2.append(["Year", "Winner", "Winner Num. of Wins", "Loser", "Loser Num. of Wins"])
    
    # Find winner and loser per year
    yearly_stats = {}
    
    
    for data in all_data:

        year = data['year']
        

        if year not in yearly_stats:
            yearly_stats[year]= {'winner': data, 'loser': data}
        else:
            if data['wins'] > yearly_stats[year]['winner']['wins']:
                yearly_stats[year]['winner'] = data
            if data['wins'] < yearly_stats[year]['loser']['wins']:
                yearly_stats[year]['loser'] = data
    
    
    for year ,stats in yearly_stats.items():

        if year < '1992':
        
            winner = stats['winner']
            loser = stats['loser']
            sheet2.append([year, winner['team'], winner['wins'], loser['team'], loser['wins']])
    
    # Save the Excel file
    wb.save(excel_filename)

# Main function to run the program
def main():
    base_url = "https://www.scrapethissite.com/pages/forms/"
    
    # Step 1: Scrape all pages and collect HTML files and data
    all_data, html_files = scrape_all_pages(base_url)
    
    # Step 2: Save HTML files to a ZIP archive
    zip_filename = "hockey_team_stats.zip"
    save_to_zip(html_files, zip_filename)
    print(f"Saved HTML files to {zip_filename}")
    
    # Step 3: Create Excel file with the scraped data
    excel_filename = "hockey_team_stats.xlsx"
    create_excel(all_data, excel_filename)
    print(f"Saved Excel file to {excel_filename}")


if __name__ == "__main__":
    main()  