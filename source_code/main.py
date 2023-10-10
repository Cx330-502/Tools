from functions import *
from models import *
from datetime import datetime


if __name__ == '__main__':
    get_times()
    get_dates()
    get_people()
    get_table()
    get_free_time()
    print("preparations over ; Consuming time :" + str(datetime.now() - time))
    time = datetime.now()
    schedule1()
    print("schedule1 over ; Consuming time :" + str(datetime.now() - time))
    time = datetime.now()
    schedule2(0, 0)
    print("schedule2 over ; Consuming time :" + str(datetime.now() - time))
    time = datetime.now()
    schedule3(0)
    print("schedule3 over ; Consuming time :" + str(datetime.now() - time))
    time = datetime.now()
    schedule4()
    print("schedule4 over ; Consuming time :" + str(datetime.now() - time))
    time = datetime.now()
    schedule5()
    print("schedule5 over ; Consuming time :" + str(datetime.now() - time))
    time = datetime.now()
    export_sche()
    print("export1 over ; Consuming time :" + str(datetime.now() - time))
    time = datetime.now()
    export_people()
    print("export2 over ; Consuming time :" + str(datetime.now() - time))
    time = datetime.now()
    for item1 in table:
        for item2 in item1['times']:
            keys = []
            for key, value in item2['people'].items():
                keys.append(key)
            print(item1['date'], item2['time'], keys)

    for people in people_0:
        print(people.name, people.status, len(people.sche_time), people.sche_time)
