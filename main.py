import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlsxwriter

baseurl = 'https://www.roco.cc/'

headers = {
    'Accept-Encoding': 'gzip, deflate, sdch',
    'Accept-Language': 'en-US,en;q=0.8',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
}

productlinks = []

for x in range(1, 2):
    r = requests.get(
        f'https://www.roco.cc/ren/products/locomotives/steam-locomotives.html?p={x}&verfuegbarkeit_status=41%2C42%2C43%2C45%2C44')
    soup = BeautifulSoup(r.content, 'lxml')
    productlist = soup.find_all('li', class_='item product product-item')

    for item in productlist:
        for link in item.find_all('a', class_='product-item-link', href=True):
            productlinks.append(link['href'])


loco_list = []
spare_parts_list = []

for link in productlinks:
    r = requests.get(link, allow_redirects=False)
    soup = BeautifulSoup(r.content, 'lxml')
    manufacturer_name = 'Roco'

    try:
        reference = soup.find('span', class_='product-head-artNr').text.strip()
    except:
        reference = ''

    try:
        price = soup.find('div', class_='product-head-price').text.strip()
    except:
        price = ''

    try:
        type = soup.find(
            'div', class_='product-head-name').h1.text.strip()
    except:
        type = ''

    try:
        scale = soup.find('td', {'data-th': 'Scale'}).text.strip()
    except:
        scale = ''

    try:
        current = soup.find('td', {'data-th': 'Control'}).text.split(' ')[0]
    except:
        current = ''

    try:
        control = soup.find('td', {'data-th': 'Control'}).text.strip()
    except:
        control = ''

    try:
        interface = soup.find('td', {'data-th': 'Interface'}).text.strip()
    except:
        interface = ''

    try:
        digital_decoder = soup.find(
            'td', {'data-th': 'Digital decoder'}).text.strip()
    except:
        digital_decoder = ''

    try:
        decoder_Type = soup.find(
            'td', {'data-th': 'Decoder-Type'}).text.strip()
    except:
        decoder_Type = ''

    try:
        motor = soup.find('td', {'data-th': 'Motor'}).text.strip()
    except:
        motor = ''

    try:
        flywheel = soup.find('td', {'data-th': 'Flywheel'}).text.strip()
    except:
        flywheel = ''

    try:
        minimum_radius = soup.find(
            'td', {'data-th': 'Minimum radius'}).text.strip()
    except:
        minimum_radius = ''

    try:
        length_over_buffer = soup.find(
            'td', {'data-th': 'Length over buffer'}).text.strip()
    except:
        length_over_buffer = ''

    try:
        number_of_driven_axles = soup.find(
            'td', {'data-th': 'Number of  driven axles'}).text.strip()
    except:
        number_of_driven_axles = ''

    try:
        number_of_axles_with_traction_tyres = soup.find(
            'td', {'data-th': 'Number of  axles with traction tyres'}).text.strip()
    except:
        number_of_axles_with_traction_tyres = ''

    try:
        coupling = soup.find('td', {'data-th': 'Coupling'}).text.strip()
    except:
        coupling = ''

    try:
        LED_lighting = soup.find(
            'td', {'data-th': 'LED lighting'}).text.strip()
    except:
        LED_lighting = ''

    try:
        head_light = soup.find('td', {'data-th': 'Head light'}).text.strip()
    except:
        head_light = ''

    try:
        LED_head_light = soup.find(
            'td', {'data-th': 'LED head light'}).text.strip()
    except:
        LED_head_light = ''

    try:
        country = soup.find(
            'td', {'data-th': 'Original (country)'}).text.strip()
    except:
        country = ''

    try:
        railway_company = soup.find(
            'td', {'data-th': 'Railway Company'}).text.strip()
    except:
        railway_company = ''

    try:
        epoch = soup.find('td', {'data-th': 'epoch'}).text.strip()
    except:
        epoch = ''

    try:
        description = soup.find(
            'div', class_='product-add-form-text').text.strip()
    except:
        description = ''

    Locomotives = {
        'Manufacturer_name': manufacturer_name,
        'Reference': reference,
        'Price': price,
        'Type': type,
        'Scale': scale,
        'Current': current,
        'Control': control,
        'Interface': interface,
        'Digital_decoder': digital_decoder,
        'Decoder_Type': decoder_Type,
        'Motor': motor,
        'Flywheel': flywheel,
        'Minimum_radius': minimum_radius,
        'Length_over_buffer': length_over_buffer,
        'Number_of_driven_axles': number_of_driven_axles,
        'Number_of_axles_with_traction_tyres': number_of_axles_with_traction_tyres,
        'Coupling': coupling,
        'LED_lighting': LED_lighting,
        'Head_light': head_light,
        'LED_head_light': LED_head_light,
        'Country': country,
        'Railway_company': railway_company,
        'Epoch': epoch,
        'Description': description,
    }

    loco_list.append(Locomotives)


df1 = pd.DataFrame(loco_list)
# # df2 = pd.DataFrame(spare_parts_list)
# # # # df3 = pd.DataFrame()
# # # # df4 = pd.DataFrame()
writer = pd.ExcelWriter('Roco - locomotives.xlsx', engine='xlsxwriter')
df1.to_excel(writer, sheet_name='Model')
# df2.to_excel(writer, sheet_name='Spare parts')
# # # # df3.to_excel(writer, sheet_name='Documents')
# # # # df4.to_excel(writer, sheet_name='Photos')
writer.save()

print('Saved to file')
