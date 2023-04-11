import PySimpleGUI as sg
import PIL.Image
import io
import base64
import requests
import json
import time
import os, pdb
import whois
import excel
from urllib.parse import urlparse
from io import BytesIO

global_data = {
    'whois': None
}

def api_call(url=None):
    """
    Will connect with urlscan.io API and retrieve an image (bytes) of provided url.
    :param url: (string) a string url of scanned website with or without https://www.
    :return: (bytes) image view in bytes object
    """
    with open('config.txt', 'r') as config:
        api_key = config.readline()[:-1]

    headers = {'API-Key': api_key, 'Content-Type': 'application/json'}
    data = {"url": url, "visibility": "private"}
    response = requests.post('https://urlscan.io/api/v1/scan/', headers=headers, data=json.dumps(data))
    print(response)
    uuid = response.json()['uuid']

    time.sleep(15)

    response = requests.get('https://urlscan.io/screenshots/' + uuid + '.png', headers=headers, data=json.dumps(data))
    print(response)
    img_bytes = response.content

    return img_bytes


def convert_to_bytes(file_or_bytes, resize=None):
    """
    Will convert into bytes and optionally resize an image that is a file or a base64 bytes object.
    Turns into  PNG format in the process so that can be displayed by tkinter
    :param file_or_bytes: either a string filename or a bytes base64 image object
    :type file_or_bytes:  (Union[str, bytes])
    :param resize:  optional new size
    :type resize: (Tuple[int, int] or None)
    :return: (bytes) a byte-string object
    :rtype: (bytes)
    """
    if isinstance(file_or_bytes, str):
        img = PIL.Image.open(file_or_bytes)
    else:
        try:
            img = PIL.Image.open(io.BytesIO(base64.b64decode(file_or_bytes)))
        except Exception:
            dataBytesIO = io.BytesIO(file_or_bytes)
            img = PIL.Image.open(dataBytesIO)

    cur_width, cur_height = img.size
    if resize:
        new_width, new_height = resize
        scale = min(new_height / cur_height, new_width / cur_width)
        img = img.resize((int(cur_width * scale), int(cur_height * scale)), PIL.Image.LANCZOS)
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    del img
    return bio.getvalue()


def show_image(data=None):
    """
    Show an image in full size in system image viewer.
    :param data: (bytes) image from urlscan.io API as a bytes object
    :return: None, it calls an im.show() and show an image
    """
    im = PIL.Image.open(BytesIO(data))
    im.show()

def save_image(url=None, path=None, data=None):
    """
    Will convert image from bytes object to image and save it as a PNG.
    :param url: (string) a string url of scanned website with or without https://www.
    :param path: (string) a string path to the place where the image should be saved
    :param data: (bytes) image from urlscan.io API as a bytes object
    :return: None, it saved an image
    """
    # Load image from BytesIO
    image_to_save = PIL.Image.open(BytesIO(data))
    image_to_save.save(path + url + '.png')

def save_whois(text=None, path=None):
    """
    Will save the passed raw text as .txt file
    :param text (string) a string containing Whois info
    :param path: (string) a string path to the place where the image should be saved
    :return: True, if it saved a text file, otherwise False
    """
    if (len(text) > 0):
        with open('{0}whois.txt'.format(path), 'a') as f:
            f.write(text)
            return True
    return False

layout = [
          [sg.Text('Save results to: '), 
            sg.InputText('C:\\', key='-PATH-'), 
            sg.Button(button_text='', image_filename='./img/folder.png', image_size=(16,16), border_width=0, button_color="#fff on #ccc", tooltip='Set to current directory', key='-SET-CURRENT-DIR-')],
          [sg.Text('URL, e.g. https://domain.to.scan.url :')],
          [sg.InputText('', key='-URL-'),
           sg.Button('Scan', key='-SCAN-', bind_return_key=True), sg.Button('Show original image size', key='-SHOW-'),
           sg.Button('Save original image', key='-SAVE-')], 
          [sg.Button('Get Whois', key='-GET-WHOIS-'), sg.Button('Save Whois to whois.txt', key='-SAVE-WHOIS-')], 
          [sg.Text('...', key='-WHOIS-RAW-')],
          [sg.Image(size=(800, 600), key='-IMAGE-')],
          [sg.Exit()],
          ]

window = sg.Window('URLSCAN API CONNECT', layout)

while True:
    event, values = window.read()
    print(event, values)  # for debugging

    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if event == '-SCAN-': # Get image from URLSCAN
        url = values['-URL-']
        try:
            image_data = api_call(url)
        except KeyError:
            sg.popup(f'The website {url} \nis currently not available. \nCheck if provided URL is correct and do not '
                     f'violate urlscan.io regulations.\n Also, check if API Key is provided and is correct.',
                     title='Error!')
            continue
        im = convert_to_bytes(image_data, resize=(800, 600))
        window['-IMAGE-'].update(data=im)

    if event == '-SHOW-': # Show image in GUI
        try:
            show_image(image_data)
        except NameError:
            sg.popup('There is no image to show. \nPlease do the scan first or scan a available website!',
                     title='Error!')
            continue

    if event == '-SAVE-': # Save image to disk
        try:
            save_image(url=url, path=values['-PATH-'], data=image_data)
        except (FileNotFoundError, NameError):
            sg.popup(f'The path is broken. \nCheck if provided PATH is correct.\nOr do the scan before saving '
                     f'the original image!', title='Error!')
            continue

    if event == '-SET-CURRENT-DIR-': # Fill in directory field with current directory
        window['-PATH-'].update(os.getcwd() + os.sep)

    if event == '-GET-WHOIS-': # Query RiskIQ and show result excerpt in Gui
        url = values['-URL-']
        parsed = urlparse(url)

        if (len(parsed.scheme) == 0 and len(parsed.netloc) == 0):
            host = parsed.path
        elif (len(parsed.netloc) > 0):
            host = parsed.netloc
        else:
            sg.popup(f'Invalid URL', title='Error!')
            pdb.set_trace()
            continue

        whois_data = whois.init_analyzer(host).whois
        global_data['whois'] = whois_data
        show_in_gui_whois = whois_data.raw['rawText']

        if (len(show_in_gui_whois.split('\n')) > 7):
            show_in_gui_whois = '\n'.join(show_in_gui_whois.split('\n')[0:6]) + '\n... [truncated]'

        window['-WHOIS-RAW-'].update(show_in_gui_whois)

    if event == '-SAVE-WHOIS-': # Save Whois result to disk
        try:
            save_whois(text=global_data['whois'].raw['rawText'], path=values['-PATH-'])
            sg.popup(f'Saved!', title='Success!')
        except (FileNotFoundError, NameError):
            sg.popup(f'The path is broken. \nCheck if provided PATH is correct.\nOr do the scan before saving '
                     f'the whois.txt file!', title='Error!')
            continue

    if event == '-SAVE-TO-EXCEL-': # Save all queried data to a new excel file


window.close()
