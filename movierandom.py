from openpyxl import Workbook, load_workbook
import random

wb = load_workbook('watchlistjjn.xlsx')
ws = wb.active

def pick_a_number(list):
    random_number_index = random.randint(2, (len(list['B'])))
    return random_number_index


def search_for_movie(movie_list, index):
    return movie_list[index]


def get_movie_list(list):
    movie_list = []
    for i in range(2, (len(list['B']))):
        movie_list.append(list['B'+str(i)].value)
    return movie_list


def get_movie_year(movie_index, list):
    movie_year = (list['C'+str(movie_index + 2)].value)
    return movie_year


def get_ltbxd_link(movie_index, list):
    movie_link = (list['D'+str(movie_index + 2)].value)
    return movie_link

        

def show_movie(movie, movie_year, movie_link):
    print('Your random movie is', movie, "(" , movie_year , "),", "check what people are saying about it in:", movie_link)


def main(list):
    try:
        number_of_movies = int(input('how many movies would you like to be recommended?'))
        if number_of_movies <= 0:
            number_of_movies = print('please select a number greater than 0, select a new number')
            main(list)
        elif number_of_movies > 0:
            
            for i in range(number_of_movies):
                movie_list = get_movie_list(list)
                movie_index = pick_a_number(list)
                movie_year = get_movie_year(movie_index, list)
                movie_link = get_ltbxd_link(movie_index, list)
                your_random_movie = search_for_movie(movie_list, movie_index)
                show_movie(your_random_movie, movie_year, movie_link)
    except ValueError:
        print('please input a number not a word')
        main(list)


main(ws)