import os


def date():
    date=str(os.popen('date "+%d/%m/%Y"').read())
    date=date.split("/")
    date[2]=date[2].split("\n")[0]
    settedDate=[]
    for i in date:
        settedDate.append(int(i))
    return settedDate

print(date())
