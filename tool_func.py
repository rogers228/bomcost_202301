from time import gmtime
from time import strftime

def fmt_seconds(seconds):
    return strftime("%H:%M:%S", gmtime(int(seconds)))

def test1():
    print(fmt_seconds(13080))

if __name__ == '__main__':
    test1()