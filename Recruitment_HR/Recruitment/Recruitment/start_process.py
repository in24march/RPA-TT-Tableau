from get_recruit import *
from map_recruit import *
from Master_HRCC import *
import time

if __name__ == '__main__':
    # date_get = get_data_web(datetime.now() - timedelta(days=1))
    date_get = datetime.now() - timedelta(days=1)
    time.sleep(1)
    run_map(date_get)   
    run_master(date_get)