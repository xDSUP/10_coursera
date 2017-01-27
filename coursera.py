import requests
import re

from random import sample
from io import BytesIO
from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_courses_list(html, courses_count=20):
    tree = etree.parse(BytesIO(html))
    root = tree.getroot()
    courses_urls = sample([url[0].text for url in root], courses_count)
    
    courses_list = []
    for url in courses_urls:
        course = get_course_info(url)
        if course is not None:
            courses_list.append(course)
    return courses_list


def get_course_info(course_slug):
    course_html = requests.get(course_slug).text
    soup = BeautifulSoup(course_html, "lxml")
    number_weeks = 0
    for week in soup.findAll(class_="week"):
        number_weeks += 1
    
    course_info = {}
    course_info['title'] = soup.find(class_='title display-3-text').text
    course_info['language'] = soup.find(class_='language-info').text
    course_info['week'] = number_weeks
    course_info['course_url'] = course_slug
    course_info['starts'] = get_starts(soup)
    course_info['rating'] = get_rating(soup)
    return course_info


def get_starts(soup):
    start_date = soup.find(class_='startdate rc-StartDateString caption-text').text
    return start_date
    

def get_rating(soup):
    rating = None
    result_tag = soup.find(class_='ratings-text bt3-hidden-xs')
    if result_tag:
        rating = re.search(r"\d+.\d+", result_tag.text).group(0)
    return rating   
    

def output_courses_info_to_xlsx(filepath, courses):
    wb = Workbook()
    ws = wb.active
    keys = ['title', 'starts', 'language', 'week', 'rating', 'course_url']

    for column, key in enumerate(keys, start=1):
        ws.cell(row=1, column=column, value=key)

    for row, item in enumerate(courses, start=2):
        for column, key in enumerate(keys, start=1):
            ws.cell(row=row, column=column, value=item[key])

    wb.save(filepath)


if __name__ == '__main__':
    website_address = "https://www.coursera.org/sitemap~www~courses.xml"
    html = requests.get(website_address).content
    courses_list = get_courses_list(html)
    filepath = input("Enter filepath (.xlsx): ")
    output_courses_info_to_xlsx(filepath, courses_list)
    print("Finish")
    
