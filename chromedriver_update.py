import os
import sys
import psutil
import zipfile
import requests
from win32com.client import Dispatch

class DownloadChrome(object):

    def __init__(self, driver_path):
        # new created variable will work after reboot the machine
        self._chrome_path = os.environ["Chrome"]
        # hack, the new set variable will work after reboot
        # self._chrome_path = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
        self._chrome_driver_path = driver_path
        self._chromedriver_url = 'https://chromedriver.storage.googleapis.com/'
        self._platform = sys.platform

    @property
    def _get_current_chrome_ver(self):
        parser = Dispatch("Scripting.FileSystemObject")
        try:
            version = parser.GetFileVersion(self._chrome_path)
        except Exception as e:
            return f'Get Chrome current version failed: {e}'
        return version

    @property
    def _get_current_chromedriver_ver(self):
        # popen must be closed after used
        curr_pipe =  os.popen(self._chrome_driver_path + ' --version')
        curr_ver = curr_pipe.read().split(' ')[1]
        curr_pipe.close()
        return curr_ver

    def _shut_down_current_driver(self):
        pids = psutil.pids()
        for pid in pids:
            if not psutil.pid_exists(pid):
                continue
            p = psutil.Process(pid)
            if p.name() == 'chromedriver.exe' and p.exe() == self._chrome_driver_path:
                print(f'kill process: {p.name()}')
                p.terminate()

    def _download_extract_driver(self, driver_url):
        with open('download_driver.zip', 'wb') as f:
            zip_file = requests.get(driver_url).content
            f.write(zip_file)
            zip_file.close()
        
        zfile = zipfile.ZipFile('download_driver.zip')
        for zf in zfile.namelist():
            if zf == 'chromedriver.exe':
                zfile.extract(zf, os.path.dirname(self._chrome_driver_path))

        zfile.close()
        os.remove('download_driver.zip')

    def _download_driver(self, chrome_ver):
        self._shut_down_current_driver()
        target_chromedriver_ver = requests.get(self._chromedriver_url + 'LATEST_RELEASE_' + str(chrome_ver.split('.')[0])).text
        chromedriver_path = self._chromedriver_url + target_chromedriver_ver + '/'
        if self._platform == 'win32':
            self._download_extract_driver(chromedriver_path + 'chromedriver_win32.zip')
        # python2 returns 'linux2'
        elif self._platform == 'linux':
            self._download_extract_driver(chromedriver_path + 'chromedriver_linux64.zip')
        else:
            self._download_extract_driver(chromedriver_path + 'chromedriver_mac64.zip')
    
    def compare_download(self):
        chrome_ver = self._get_current_chrome_ver
        chromedriver_ver = self._get_current_chromedriver_ver
        if chrome_ver.split('.')[0] != chromedriver_ver.split('.')[0]:
            self._download_driver(chrome_ver)        
        return f'Chrome: {chrome_ver}\nChromeDriver: {self._get_current_chromedriver_ver}'

if __name__ == "__main__":
    DownloadChrome(os.path.join(os.path.split(os.path.abspath(__file__))[0], 'chromedriver.exe')).compare_download()
