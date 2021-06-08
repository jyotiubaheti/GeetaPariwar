from os import walk
import pandas
from datetime import datetime
import re
import collections
from functools import reduce
import sys


def getDigits(str1):
    return list(set(re.findall(r'\d+', str1)))

def zoomToTime(start):
    print (start)
    day = start.split(" ")[0]
    time = start.split(" ")[1].split(":")
    if start.split(" ")[-1] == "PM":
        time[0] = str(int(time[0]) + 12)
    if len(time) == 2:
        time.append("00")
    time = ":".join(time)
    if "/" in day:
        obj = datetime.strptime(day + " " + time, '%m/%d/%Y %H:%M:%S')
    elif "-" in day:
        obj = datetime.strptime(day + " " + time, '%m-%d-%Y %H:%M:%S')
    else:
        exit("time format not recognized")
    return obj

# get files
f = []
for (dirpath, dirnames, filenames) in walk("./"):
    for file in filenames:
        if ".csv" in file:
            f.append(file)
        #end:
    break
    #end:
#end:

print (f)
dfDict = {}
tagStart = 'Start Time'
tagEnd = 'End Time'
tagName = 'Name (Original Name)'
tagRollNumber = "Roll Number"
tagTime = 'Total Duration (Minutes)'
for file in f:
    df = pandas.read_csv(file)
    start = zoomToTime(df[tagStart][0])
    end = zoomToTime(df[tagEnd][0])

    i=0
    while 1:
        if df.iloc[i][0] == tagName:
            break
        i += 1

    df.columns = df.iloc[i]
    df = df.iloc[i+1:]
    df[tagRollNumber] = df[tagName].apply(getDigits)
    df = df.explode(tagRollNumber)
    df[tagTime] = df[tagTime].fillna(0).astype(int)
    df[tagRollNumber] = df[tagRollNumber].fillna(0).astype(float)
    df = df.groupby([tagRollNumber])[tagTime].agg('sum').reset_index()
    df.columns = [tagRollNumber,start]

    dfDict[(start - datetime.fromtimestamp(0)).total_seconds()] = df

od = collections.OrderedDict(sorted(dfDict.items()))

finalDF = reduce(lambda  left,right: pandas.merge(left,right,on=[tagRollNumber],
                                            how='outer'), od.values())

finalDF = finalDF.sort_values(tagRollNumber)
finalDF = finalDF.append(finalDF.count().rename('Total'), ignore_index=True)
# print (finalDF)
finalDF.to_excel("zoomAttendance.xlsx",index=False)