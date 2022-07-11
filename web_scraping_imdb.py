import xlsxwriter
import requests
import bs4

URL = "https://www.imdb.com/chart/top/?ref_=nv_mv_250"

result = requests.get(URL)
request = bs4.BeautifulSoup(result.text, 'html.parser')
filme = request.find('tbody', class_='lister-list').find_all('tr')

# Am stocat in variabile tipul de data pe care il vreau.
context = {'data': []}
for film in filme:
    data = {}
    nume_film = film.find('td', class_='titleColumn').a.text
    rank = film.find('td', class_='titleColumn').get_text(strip=True).split('.')[0]
    an = film.find('td', class_='titleColumn').span.text.strip('()')
    rating = film.find('td', class_='ratingColumn imdbRating').strong
    rating_to_string = str(rating)
    descriere = rating_to_string[15:-14]

    # tratez cazurile in care nu gaseste
    if rating:
        data['rank'] = rank
    else:
        data['rank'] = 'No data available'

    if nume_film:
        data['nume_film'] = nume_film
    else:
        data['nume_film'] = 'Niciun film gasit'

    if an:
        data['an'] = an
    else:
        data['an'] = 'No data available'

    if descriere:
        data['descriere'] = descriere
    else:
        data['descriere'] = 'No data available'

    context['data'].append(data)

workbook = xlsxwriter.Workbook('Lista_Filme.xlsx')
worksheet = workbook.add_worksheet('Lista_filme')

rand = 0
coloana = 0

for film in context['data']:

    worksheet.write(rand, coloana, film['rank'])
    worksheet.write(rand, coloana + 1, film['nume_film'])
    worksheet.write(rand, coloana + 2, film['an'])
    worksheet.write(rand, coloana + 3, film['descriere'])
    rand += 1

workbook.close()
