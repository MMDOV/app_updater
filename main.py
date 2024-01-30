import subprocess
import re
import time
import os
from bs4 import BeautifulSoup
import requests
import lxml
import webbrowser
import win32api

FILES_PATH = {
    "IDM": r"C:\Program Files (x86)\Internet Download Manager\IDMan.exe"
}


def get_driver_version():
    """Get the installed version of nvidia driver"""
    cmd = r'reg query "HKEY_LOCAL_MACHINE\SOFTWARE\NVIDIA Corporation\Installer2\Stripped" /s | find "Display.Driver/"'
    output = subprocess.check_output(cmd, shell=True)
    all_ = [float(x) for x in re.findall(r'Display\.Driver/(\d+\.?\d*)', str(output))]
    latest_version_installed = max(all_)
    return latest_version_installed


def get_file_properties(file_name):
    """
    Read all properties of the given file return them as a dictionary.
    """
    prop_names = ('Comments', 'InternalName', 'ProductName',
                  'CompanyName', 'LegalCopyright', 'ProductVersion',
                  'FileDescription', 'LegalTrademarks', 'PrivateBuild',
                  'FileVersion', 'OriginalFilename', 'SpecialBuild')

    props = {'FixedFileInfo': None, 'StringFileInfo': None, 'FileVersion': None}

    try:
        # backslash as parm returns dictionary of numeric info corresponding to VS_FIXEDFILEINFO struc
        fixed_info = win32api.GetFileVersionInfo(file_name, '\\')
        props['FixedFileInfo'] = fixed_info
        props['FileVersion'] = "%d.%d.%d.%d" % (fixed_info['FileVersionMS'] / 65536,
                                                fixed_info['FileVersionMS'] % 65536,
                                                fixed_info['FileVersionLS'] / 65536,
                                                fixed_info['FileVersionLS'] % 65536)

        # \VarFileInfo\Translation returns list of available (language, codepage)
        # pairs that can be used to retrieve string info. We are using only the first pair.
        lang, codepage = win32api.GetFileVersionInfo(file_name, '\\VarFileInfo\\Translation')[0]

        # any other must be of the form \StringfileInfo\%04X%04X\parm_name, middle
        # two are language/codepage pair returned from above

        str_info = {}
        for prop_name in prop_names:
            str_info_path = u'\\StringFileInfo\\%04X%04X\\%s' % (lang, codepage, prop_name)
            ## print str_info
            str_info[prop_name] = win32api.GetFileVersionInfo(file_name, str_info_path)

        props['StringFileInfo'] = str_info
    except:
        pass

    return props


def get_latest_driver():
    """Get the latest nvidia driver version available on Yas-dl"""
    site_ad = r'https://www.yasdl.com/33655/%D8%AF%D8%A7%D9%86%D9%84%D9%88%D8%AF-%D8%AF%D8%B1%D8%A7%DB%8C%D9%88%D8%B1' \
              r'-nvidia-geforce.html'
    response = requests.get(site_ad)
    text = response.text
    soup = BeautifulSoup(text, "lxml")
    post_title = soup.find(name="h1", class_="post-title").getText()
    latest_ver = float(re.findall(r"[-+]?(?:\d*\.*\d+)", post_title)[0])
    return latest_ver, soup


def get_latest_app_ver(site_address):
    """Get the latest App version available on soft98.ir"""
    headers = {
        "Accept-Language": "en-US,en;q=0.5",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Referer": "https://soft98.ir/tags/download+manager/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/117.0",
    }
    response = requests.get(site_address, headers=headers)
    text = response.text
    soup = BeautifulSoup(text, "lxml")
    post_title = soup.find(name="a", class_="cbddtl").get_text()
    latest_ver = re.findall(r"[-+]?(?:\d*\.*\d+\.*\d+)", post_title)[0]
    return latest_ver, soup


def get_app_link(app_version, latest_app_version, bfs):
    """If there is a newer version of App get the link and open it in browser"""
    app_version_list = app_version.split('.')
    latest_app_version_list = latest_app_version.split('.')
    for i in range(len(latest_app_version_list)):
        try:
            if int(latest_app_version_list[i]) > int(app_version_list[i]):
                print("Updating", end='')
                time.sleep(0.5)
                print(".", end='')
                time.sleep(0.5)
                print(".", end='')
                time.sleep(0.5)
                print(".", end='')
                dl_box = bfs.find(name="dl", id="dbdll")
                dl_link = dl_box.find(name="a", class_="dbdlll")['href']
                webbrowser.open(dl_link)
        except IndexError:
            print("Updating", end='')
            time.sleep(0.5)
            print(".", end='')
            time.sleep(0.5)
            print(".", end='')
            time.sleep(0.5)
            print(".", end='')
            dl_box = bfs.find(name="dl", id="dbdll")
            dl_link = dl_box.find(name="a", class_="dbdlll")['href']
            webbrowser.open(dl_link)


def get_nvidia_driver_link(latest_driver_installed, latest_ver, bfs):
    """If there is a newer version of nvidia driver get the download link and open it in browser"""
    if latest_driver_installed < latest_ver:
        print("Updating", end='')
        time.sleep(0.5)
        print(".", end='')
        time.sleep(0.5)
        print(".", end='')
        time.sleep(0.5)
        print(".", end='')
        dl_box = bfs.find(name="ul", id="dl-box1")
        dl_link = dl_box.find(name="a", class_="col")["href"]
        webbrowser.open(dl_link)
    else:
        print("The latest version is already installed!")


# Get what app user wants to update
user_choice = int(input(
    "What App Do You Want To Update?\n"
    "-----------------------------------\n"
    "1 -- Nvidia Driver \n"
    "2 -- IDM \n"
    "3 -- Pycharm \n"
    "-----------------------------------\n"))
if user_choice == 1:
    latest_installed = get_driver_version()
    latest_version, soup = get_latest_driver()
    get_nvidia_driver_link(latest_driver_installed=latest_installed, latest_ver=latest_version, bfs=soup)
elif user_choice == 2:
    idm_info = get_file_properties(FILES_PATH.get('IDM'))
    idm_installed = idm_info.get("FileVersion")
    latest_version, soup = get_latest_app_ver(
        r"https://soft98.ir/internet/download-manager/4-internet-download-manager-4.html")
    get_app_link(app_version=idm_installed, latest_app_version=latest_version, bfs=soup)
elif user_choice == 3:
    # Get latest installed PyCharm version
    all_installed_apps = [name for name in os.listdir(r"C:\Program Files\JetBrains")]
    all_versions = [x.strip('PyCharm ') for x in all_installed_apps]
    latest_installed_version = max(all_versions)
    # Get latest version available
    latest_version, soup = get_latest_app_ver(r"https://soft98.ir/software/programming/1652-pycharm.html")
    # Get and open the link if there is a newer version
    get_app_link(app_version=latest_installed_version, latest_app_version=latest_version, bfs=soup)
