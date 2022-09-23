from time import gmtime
from time import strftime
import math

def fmt_seconds(seconds):
    return strftime("%H:%M:%S", gmtime(int(seconds)))

def is_dec(number):
    # 是否有小數
    if any([type(number)==int, type(number)==float]):
        return number-math.floor(number) >= 0.001
    else:
        return False

def test1():
    print(fmt_seconds(13080))



if __name__ == '__main__':
    test1()