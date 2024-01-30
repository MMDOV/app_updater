import subprocess
import re
import time
from bs4 import BeautifulSoup
import requests
import lxml
import webbrowser
from tkinter import messagebox

# Get the installed version of driver
cmd = r'reg query "HKEY_LOCAL_MACHINE\SOFTWARE\NVIDIA Corporation\Installer2\Stripped" /s | find "Display.Driver/"'
output = subprocess.check_output(cmd, shell=True)
all_ = [float(x) for x in re.findall(r'Display\.Driver/(\d+\.?\d*)', str(output))]
latest_installed = max(all_)


def get_latest():
    # Get the latest version available on Yas-dl
    site_ad = r'https://www.yasdl.com/33655/%D8%AF%D8%A7%D9%86%D9%84%D9%88%D8%AF-%D8%AF%D8%B1%D8%A7%DB%8C%D9%88%D8%B1' \
              r'-nvidia-geforce.html'
    response = requests.get(site_ad)
    text = response.text

    soup = BeautifulSoup(text, "lxml")
    print(soup)
    post_title = soup.find(name="h1", class_="post-title").getText()
    latest_ver = float(re.findall(r"[-+]?(?:\d*\.*\d+)", post_title)[0])

    # If there is a newer version get the download link
    if latest_installed < latest_ver:
        dl_box = soup.find(name="ul", id="dl-box1")
        dl_link = dl_box.find(name="a", class_="col")["href"]
        webbrowser.open(dl_link)
    else:
        messagebox.showinfo("The latest version is already installed!")


try:
    get_latest()
except TypeError:
    webbrowser.open(r'https://www.yasdl.com/')
    time.sleep(3)
    get_latest()
