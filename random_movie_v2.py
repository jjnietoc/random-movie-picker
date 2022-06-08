from openpyxl import Workbook, load_workbook
import random

wb = load_workbook('watchlistjjn.xlsx')
ws = wb.active

class LbxMovieSource: 
    def __init__(self, source):
        self.source = source

    def createMovList(self, source): 
        movie_list = []
        for element in range(2, (len(source['B']))):
            movie_list.append(source['B'+str(element)].value)
        return movie_list  

    def createRelList(self, source):
        release_list = []
        for element in range(2, (len(source['C']))):
            release_list.append(source['C'+str(element)].value)
        return release_list

    def createAddList(self, source):
        added_list = []
        for element in range(2, (len(source['A']))):
            added_list.append(source['A'+str(element)].value)
        return added_list

    def createUrlList(self, source):
        url_list = []
        for element in range(2, (len(source['D']))):
            url_list.append(source['D'+str(element)].value)
        return url_list

class listWorker:
    def __init__(self, movie_list, release_list, added_list, url_list):
        self.movie_list = movie_list
        self.release_list = release_list
        self.added_list = added_list
        self.url_list = url_list

    def joinLists(self): #medio porlas si ya existe joinAndSlice
        join_list = [val for pair in zip(movie_list, release_list, added_list, url_list) for val in pair]
        return join_list

    def joinAndSlice(self):
        join_list = [val for pair in zip(movie_list, release_list, added_list, url_list) for val in pair]
        movies = [join_list[x:x+4]for x in range(0, len(join_list), 4)]
        return movies
           
class randomMovieEngine:
    def __init__(self, source):
        self.source = source

    def randomMoviePick(self):
        randNum = random.randint(0, len(self.source))
        your_movie = self.source[randNum]
        return your_movie

class Movie:
    def __init__(self, movie_info):
        self.movie_info = movie_info

    def presentInfo(self):
        movie_name = self.movie_info[0]
        release_date = self.movie_info[1]
        added_date = self.movie_info[2]
        url = self.movie_info[3]
        finalMovie = "Your random movie is: " + str(movie_name) + " (" + str(release_date) + "). \n" + "This movie has been in your watchlist since: " + str(added_date) + ". \n" + "see what people are saying about it: " + str(url) + "."
        return finalMovie
            

#test result: todo funciona bien, tengo 4 listas de cada cosa
movies = LbxMovieSource(ws)
movie_list = movies.createMovList(ws)
release_list = movies.createRelList(ws)
added_list = movies.createAddList(ws)
url_list = movies.createUrlList(ws)

#test result 2: todo ok, tengo una lista con toda la info de cada pelicula
test = listWorker(movie_list, release_list, added_list, url_list)
lists = test.joinAndSlice()

#test result 3: todo ok, devuelve un elemento de la lista = pelicula y su info
randomizer = randomMovieEngine(lists)
test2 = randomizer.randomMoviePick()

#test final: todo ok! te da una pelicula al azar
myMovie = Movie(test2)
final_test = myMovie.presentInfo()
print(final_test)
