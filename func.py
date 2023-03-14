# for some function
import random
import time

def random_wait(t_wait, info, random_deviation = 60):
    """random wait with dynamic print 

    Args:
        t_wait (_type_): _description_
        info (_type_): _description_
    """
    t_wait = random.randint(t_wait -  random_deviation, t_wait + random_deviation)
    while t_wait:
        flush_print(f'{info} {convert_seconds_to_hms(t_wait)}')
        time.sleep(1)
        t_wait -= 1

def flush_print(info):
    print(f'\r\033[K {info}', end='')

def convert_seconds_to_hms(seconds):
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = (seconds % 3600) % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"