import requests
import base64


def png_loading_imgbb(path):
    #   Функция загружает картинку на сайт через API и возвращает ссылку на её
    url = "https://api.imgbb.com/1/upload"
    key = 'key'
    with open(path, 'rb') as file:
        encoded = base64.b64encode(file.read())
    payload = {
        "key": key,
        "image": encoded,
    }
    res = requests.post(url, payload)
    return res.json()['data']['url_viewer']
