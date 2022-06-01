from base64 import b64encode, b64decode
from binascii import hexlify
from datetime import datetime
from requests import post
from bs4 import BeautifulSoup
from time import time
from uuid import uuid4
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def b64benc(s):
    return b64encode(s).decode("utf-8")

def b64senc(s):
    return b64benc(s.encode("utf-8"))

def b64hdec(s):
    return hexlify(b64decode(s)).decode("utf-8")

def iso_date(t):
    return datetime.fromtimestamp(t).isoformat() + "Z"

def header_data():
    uuid = str(uuid4())
    create_time = time()
    create_date = iso_date(create_time)
    expire_date = iso_date(create_time + 120)
    
    return uuid, create_date, expire_date

def wu_request(url, data):
    text = post(url, headers={
        'User-Agent': 'Windows-Update-Agent/10.0.10011.16384 Client-Protocol/2.50',
        'Content-Type': 'application/soap+xml; charset=utf-8'
    }, data=data, verify=False).text
    soup = BeautifulSoup(text, features="xml")
    
    return soup

def parse_xml(xml):
    return BeautifulSoup(xml, features="lxml")
