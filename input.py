# take the input from Xl and translate it to matrix (m,n) while m is days to work and n is num of worker
# can be an argoment of file or keyboard

import pandas as pd
from distribution import worker



class inputA:
    def __init__(self, paths, pref = True): #check the path
        self.paths = paths
        self.num_of_worker = 0
        self.res = []
        self.name = ""
        self.days = []
        self.regions = []
        self.names = []
        self.list_worker = []
        self.used = False
        self.pref =pref

    def findWorker(self, name):
        for man in self.list_worker:
            if man.name == name:
                return man
        return None

    #extracting the day of workers file
    def ReadData(self):
        for filename in self.paths:
            day_to_work = []
            counter = 1
            data = pd.read_excel(filename)

            #Extracting days of work
            if not self.used:
                self.used = True
                for row in data["תאריך"]:
                    type_of_day = type(row)
                    break
                for row in data["תאריך"]:
                    temp = ""
                    if type(row ) is type_of_day:
                        for i in range(len(str(row))):
                            if (str(row)[i]) == " ":
                               break
                            temp += str(row)[i]
                        self.days.append(temp)

            # Extracting days of worker and is name
            for d in data["שם"]:
                if type(d) is str:
                    name = d
                    day_to_work.append(counter)
                counter += 1
            self.list_worker += [worker([],day_to_work, name)]  #need to update tegion of worker later

    # extracting the region and names of worker from "region.xlsx" file
    def TakeRegionOfWorker(self, path="system/regions.xlsx"):
        temp = []
        data = pd.read_excel(path, header=0)

        #Extracting the all regions
        for i in data.head(0):
            temp += [i]
        self.regions = [temp[i] for i in range(6)]


        # Extracting the regions of worker
        for i in data.values:
            is_region = []
            name  = i[len(data.values)-1]
            for j in range(len(i)):
                # if there is solution with prefenses add the X value
                if self.pref and i[j] is "X": # 7 is max day. can be upgrade
                    is_region += [j]
                # if there is not solution with prefenses add also - value
                if not self.pref and (i[j] is "X" or i[j] is "-"):
                    is_region += [j]
            if self.findWorker(name) is not None:
                self.findWorker(name).regions = is_region
            #if the worker did not fill his day take the all days
            else: self.list_worker += [worker(is_region, [i+1 for i in range(len(self.days))],name)]

        #if we get worker from files worker that are not in the "region.xlsx" return -1
        for w in self.list_worker:
            if len(w.regions) == 0:
                return -1




