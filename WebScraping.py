import requests, openpyxl
from bs4 import BeautifulSoup

# Creating a new Excel workbook and printing its sheet names
excel = openpyxl.Workbook()
print(excel.sheetnames)

#Setting the active sheet's title to 'Top Rated Movies' and printing the updated sheet names
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)

#Adding the headers for the data we want to collect
sheet.append(['Movie Rank','Movie Name','Year of Release','IMDB Rating'])


try:
     # Sending a GET request to the URL and storing the response in source
    source = requests.get('https://www.imdb.com/chart/top')

    # Raise an exception if the status code of the response is not OK
    source.raise_for_status()

    # Creating a BeautifulSoup object to parse the HTML content of the page
    soup = BeautifulSoup(source.text,'html.parser')

     # Finding the tbody element with class 'lister-list' and getting all the tr elements inside it
    movies = soup.find('tbody', class_="lister-list").find_all('tr')

    # Looping through each movie element to extract the details
    for movie in movies:
        
        name = movie.find('td', class_="titleColumn").a.text

        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]

        year = movie.find('td', class_="titleColumn").span.text.strip('()')

        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text

        print(name, rank, year, rating)
        sheet.append([name, rank, year, rating])
        

    
    

except Exception as e:
    # Catching any exceptions that may arise and printing the error message to the console
    print(e)

excel.save('IMDB Movie Ratings.xlsx')