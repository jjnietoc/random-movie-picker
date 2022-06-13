from openpyxl import Workbook, load_workbook
import random

wb = load_workbook('watchlistjjn.xlsx')
ws = wb.active

class Movie:
    def __init__(self, movie_name, release_date, date_added, url):
        self.movie_name = movie_name
        self.release_date = release_date
        self.date_added = date_added
        self.url = url

    def __str__(self):
        your_movie = "Your random movie is: " + str(self.movie_name) + " (" + str(self.release_date) + ")" + ". \n"
        since = "This movie has been in your watchlist since: " + str(self.date_added) + ". \n"
        opinions = "See what people are saying about it in: " + str(self.url) + "."
        return your_movie + since + opinions

class  MovieSource:
    def __init__(self, source):
        self.source = source

    def create_movie_object(self):
        length = len(self.source['A'])
        B = []
        for elements in range(2, length):
            B.append(self.source['B' + str(elements)].value)
        
        C = []
        for elements in range(2, length):
            C.append(self.source['C' + str(elements)].value)

        A = []
        for elements in range(2, length):
            A.append(self.source['A' + str(elements)].value)

        D = []
        for elements in range(2, length):
            D.append(self.source['D' +str(elements)].value)

        all_movies = []
        for elements in B, C, A, D:
            for item1, item2, item3, item4 in zip(B, C, A, D):
                all_movies.append(Movie(item1,item2,item3,item4))
            return all_movies

class RandomMovieEngine:
    def __init__(self, source):
        self.source = source

    def SelectMovie(self):
        random_number = random.randint(0,len(self.source))
        your_movie = self.source[random_number]
        return your_movie



test = MovieSource(ws)
test2 = test.create_movie_object()
test3 = RandomMovieEngine(test2)
test4 = test3.SelectMovie()
print(test4)

