import constraint
import numpy as np
from itertools import combinations


class worker:
    index = 0
    regions = []
    day = []
    name = ""

    # here we initialling worker with the days that he can work and his region
    # parameters - r: a kist of his regions, d: a list of his days, name: the name of the worker

    def __init__(self, r, d, name):
        self.regions = r
        self.day = d
        self.name = name

    def __str__(self):
        return self.name

class Arrangement:
    date = ""
    region = ""
    worker = ""

    # here we we define the arrangement for a worker
    # parameters - date: his date, region: list of his region, worker: his name

    def __init__(self, date, region, worker):
        self.date = date
        self.region = region
        self.worker = worker

    def __str__(self):
        #print("the worker with name", end=" ")
        #print(self.worker, end=" ")
        #print("is day to work is", end=" ")
        #print(self.date, end=" ")
        #print("and is region is", end=" ")
        #print(self.region)
        return "the worker with name "+self.worker+" is day to work is "+self.date + " and is region is "+self.region


class distribution:

    # here we initialling the class which use to callculate the optimal arrangement.
    # Taking into account the constraints of days, regions, and giving weight to the
    # number of days a worker gives, and to its relative and constant number.
    # parameters - w: list of class worker, region_name: list of the name of the all regions, date: all days of arrangement
    # weight: this param define how much we decide to div the weight of a worker which get arrangement for 1 day.

    def __init__(self, w , regions_name, date, weight = 4):
        self.workers = w
        self.regions = regions_name
        self.date = date
        self.param = weight
        self._Weight = [len(i.day) for i in w]
        for i in range(len(w)):
            self.workers[i]._index = i+1


    def is_legal(self, sol):
        day_exist =[]
        for day in sol:
            for item in day:
                if day_exist.__contains__(item[len(item)-1]):
                    return False
                day_exist += item[len(item)-1]
                break
        return True

    # calculating the value of a given day arrangement.
    # parameters - sol: list of given day arrangement.

    def _calvalue(self, sol):
        temp = self._Weight.copy()
        for res in sol:
            index_of_name = ""
            for i in range(len(res)):
                if res[i] is ",":
                    break
                index_of_name += res[i]
            temp[int(index_of_name)-1] /= self.param
        return sum(temp)

    # convert index to is region

    def _convertRegions(self,index):
        return self.regions[index]

    # convert index to a name of a worker

    def _convertNames(self,index):
        for man in self.workers:
            if man._index is index:
                return man.name

    # convert index to a date

    def _convertDay(self, index):
        return self.date[index-1]

    # here we return the final arrangement

    def _Restore(self, sol):
        if not self.is_legal(sol):
            return None
        Arrangements = []
        for day in sol:
            for res in day:
                index_of_name = ""
                for i in range(len(res)):
                    if res[i] is ",":
                        break
                    index_of_name += res[i]
                name = self._convertNames(int(index_of_name))
                is_dey = self._convertDay(int(res[len(res)-1]))
                is_region = self._convertRegions(int(day[res]))
                Arrangements += [Arrangement(is_dey, is_region,name)]
        return Arrangements

    #find optimal arrangement for a list of optional arrangement.
    # parameters - solutions: lists of list of day potential solution.

    def _find_opt(self,solutions):
        opt = []
        temp = {}
        for day in solutions:
            min = np.inf
            for sol in day:
                val = self._calvalue(sol)
                if val < min:
                    min = val
                    temp = sol
            opt.append(temp)
            for res in temp:
                index_of_name = ""
                for i in range(len(res)):
                    if res[i] is ",":
                        break
                    index_of_name += res[i]
                self._Weight[int(index_of_name) - 1] /= self.param
        return self._Restore(opt)

    # find optimal arrangement for parameters in class

    def findmatch(self):
        solutions = []
        for i in range(len(self.date)):
            temp = []
            var = []
            for man in self.workers:
                if man.day.__contains__(i+1):
                    rank = [j for j in man.regions]
                    var += [(str(man._index)+","+str(i+1), rank)]
            # getting permutation of worker.
            # for example for 10 worker and 6 regions we will check all combination of 6 from 10.
            comb = combinations(var, len(self.regions))
            for p in comb:
                problem = constraint.Problem()
                for permutation in p:
                    problem.addVariable(permutation[0], permutation[1])
                problem.addConstraint(constraint.AllDifferentConstraint())
                temp += problem.getSolutions()
            if len(temp) == 0:
                return None
            solutions += [temp]

        if len(solutions) == 0:
            return None
        return self._find_opt(solutions)


