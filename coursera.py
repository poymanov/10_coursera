import requests
import sys
import xml.etree.ElementTree as ET
from random import randint
from bs4 import BeautifulSoup
from openpyxl import Workbook
from datetime import datetime


def get_courses_list():
    courses_list = []

    feed_url = 'https://www.coursera.org/sitemap~www~courses.xml'
    response = requests.get(feed_url)

    if not response.ok:
        return None

    xml_content = response.text
    courses_data = ET.fromstring(xml_content)

    xml_namespace = '{http://www.sitemaps.org/schemas/sitemap/0.9}'

    link_elements = courses_data.findall('.//{}loc'.format(xml_namespace))
    link_elements_length = len(link_elements)

    courses_count = 1

    for __ in range(courses_count):
        link_item_number = randint(0, link_elements_length)
        link_element = link_elements[link_item_number]
        courses_list.append(link_element.text)

    return courses_list


def get_courses_info(courses_list):
    courses_info = []
    for course_url in courses_list:
        course_data = get_course_data(course_url)

        if course_data:
            courses_info.append(course_data)

    return courses_info


def get_course_data(course_url):
    response = requests.get(course_url)

    if not response.ok:
        return None

    courses_info = {}

    soup = BeautifulSoup(response.text, 'html.parser')

    title = soup.find('h1', 'title display-3-text').text

    language = soup.find('div', 'rc-Language').text

    starts = soup.find('div', {'id': 'start-date-string'}).text
    starts = starts.split()
    starts = '{} {}'.format(starts[-2], starts[-1])

    weeks = soup.find_all('div', 'week')
    weeks_count = len(weeks)

    rating = soup.find('div', 'ratings-text bt3-visible-xs')

    if rating:
        rating = rating.text
        rating = rating.split()[0]
    else:
        rating = ''

    courses_info['title'] = title
    courses_info['language'] = language
    courses_info['starts'] = starts
    courses_info['weeks'] = weeks_count
    courses_info['rating'] = rating

    return courses_info


def output_courses_info_to_xlsx(courses_info):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(['Title', 'Language', 'Stars', 'Weeks', 'Rating'])

    for course in courses_info:
        worksheet.append([course['title'], course['language'],
                          course['starts'], course['weeks'], course['rating']])

    current_date_format = datetime.today().strftime('%Y-%m-%d-%H-%M-%S')
    filename = '{}.xlsx'.format(current_date_format)
    workbook.save(filename)
    print('Courses info saved to {}'.format(filename))


if __name__ == '__main__':
    courses_list = get_courses_list()

    if not courses_list:
        sys.exit('Failed to get courses urls')

    courses_info = get_courses_info(courses_list)

    if not courses_info:
        sys.exit('Failed to get courses info')

    output_courses_info_to_xlsx(courses_info)
