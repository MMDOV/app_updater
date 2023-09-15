import subprocess
import re
import time

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
        # pairs that can be used to retreive string info. We are using only the first pair.
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


def get_latest_idm_ver():
    """Get the latest Internet Download Manager version available on soft98.ir"""
    site_address = r"https://soft98.ir/internet/download-manager/4-internet-download-manager-4.html"
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


def get_idm_link(idm_version, latest_idm_version, bfs):
    idm_version_list = idm_version.split('.')
    latest_idm_version_list = latest_idm_version.split('.')
    for i in range(len(idm_version_list)):
        try:
            if int(latest_idm_version_list[i]) > int(idm_version_list[i]):
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
            break


def get_nvidia_driver_link(latest_driver_installed, latest_ver, bfs):
    """If there is a newer version of nvidia driver get the download link"""
    if latest_driver_installed < latest_ver:
        dl_box = bfs.find(name="ul", id="dl-box1")
        dl_link = dl_box.find(name="a", class_="col")["href"]
        webbrowser.open(dl_link)
    else:
        print("The latest version is already installed!")


user_choice = int(input("What app do you want to update ?\n1 -- Nvidia Driver \n 2 -- for IDM \n"))
if user_choice == 1:
    latest_installed = get_driver_version()
    latest_version, soup = get_latest_driver()
    get_nvidia_driver_link(latest_driver_installed=latest_installed, latest_ver=latest_version, bfs=soup)
elif user_choice == 2:
    idm_info = get_file_properties(FILES_PATH.get('IDM'))
    idm_installed = idm_info.get("FileVersion")
    latest_version, soup = get_latest_idm_ver()
    get_idm_link(idm_version=idm_installed, latest_idm_version=latest_version, bfs=soup)
