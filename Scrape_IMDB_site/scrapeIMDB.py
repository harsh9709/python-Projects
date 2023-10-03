from bs4 import BeautifulSoup
import requests,re,openpyxl



excel=openpyxl.Workbook()
print(excel.sheetnames)
sheet=excel.active
sheet.title='Top 250 IMDB Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie Rank','Movie Name','Year of Release','IMDB Rating'])




# User-Agent Header: Set a user-agent header in your request to make it appear as if it's coming from a web browser, rather than a bot. Many websites block requests without a user-agent header.
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'
}

try:
    source = requests.get('https://www.imdb.com/chart/top/', headers=headers)
    source.raise_for_status()
    soup = BeautifulSoup(source.text, 'html.parser')
    movies =soup.find('ul',class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-3f13560f-0 sTTRj compact-list-view ipc-metadata-list--base").find_all('li')
    for movie in movies:
        name=' '.join(movie.find('a',class_="ipc-title-link-wrapper").h3.text.split(" ")[1:])
        rank=movie.find('a',class_="ipc-title-link-wrapper").h3.text.split('.')[0]
        year=movie.find('div',class_="sc-b51a3d33-5 ibuRZu cli-title-metadata").span.text
        rating_text=movie.find('div',class_="sc-e3e7b191-0 iKUUVe sc-b51a3d33-2 ccUQup cli-ratings-container").span.text
        rating_regex=re.search(r'(\d+\.\d+)', rating_text)
        rating=rating_regex.group(1)


        print(rank,name,year,rating)
        sheet.append([rank,name,year,rating])

except Exception as e:
    print(e)

excel.save('IMDB Movie Ratings.xlsx')