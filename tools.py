import subprocess
from time import sleep


def copyToClipboard(text: str):
    process = subprocess.Popen(
    'pbcopy', env={'LANG': 'en_US.UTF-8'}, stdin=subprocess.PIPE)
    process.communicate(text.encode('utf-8'))

def readFromClipboard() -> str:
    return subprocess.check_output(
        'pbpaste', env={'LANG': 'en_US.UTF-8'}).decode('utf-8')

def countDown(preMsg: str, postMsg: str, amount: int):
    for i in range(amount, 0, -1):
        print(preMsg + str(i) + postMsg, end = '\r')
        sleep(1)