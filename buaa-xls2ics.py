import re, xlrd, ics, arrow
import logging

v2 = r'[\,\，\s]+'
re_timesplit = re.compile(v2)


v4 = r'(\d+)[\-\－]?(\d*)'
re_repeat = re.compile(v4)

session_begin_time = arrow.get('2017-09-18T08:00:00+08:00')

xueyuanlu_time = {\
    1 : (8, 00), \
    2 : (8, 55), \
    3 : (10, 00), \
    4 : (10, 55), \
    5 : (14, 00), \
    6 : (14, 55), \
    7 : (16, 00), \
    8 : (16, 55), \
    9 : (18, 00), \
    10 : (18, 55), \
    11 : (20, 00), \
    12 : (20, 55)}

shahe_time = {
    1 : (8, 10), \
    2 : (9, 10), \
    3 : (10, 10), \
    4 : (11, 10), \
    5 : (13, 30), \
    6 : (14, 30), \
    7 : (15, 30), \
    8 : (16, 30), \
    9 : (18, 20), \
    10 : (19, 20), \
    11 : (20, 20), \
    12 : (21, 20)
}

class ClassInfoHandle:
    def __init__(self, column, row, redict):
        self.time = []
        self.time.append(int(column))
        self.class_name = redict['class_name']
        self.teacher = redict['teacher']
        self.repeat = redict['repeat']
        di_index = redict['place_and_time'].find('第')
        jie_index = redict['place_and_time'].find('节')
        if di_index:
            self.class_room = redict['place_and_time'][:di_index]
            self.time_str = redict['place_and_time'][di_index+1:jie_index]
        else:
            self.class_room = ''
            self.time_str = redict['place_and_time'][di_index+1:jie_index]
        self._time = re_timesplit.split(self.time_str)
        self._time = list(map(int, self._time))
        self.time.extend(self._time)
        self._repeat = None
        self.rrule = None


        self._repeat = list(re_repeat.match(self.repeat).groups())
        if self._repeat[-1] == '':
            self._repeat.remove('')
        self._repeat = list(map(int, self._repeat))

        if not self.repeat.find('双周') == -1:
            self._repeat.append(2)
            self._repeat[0] += 1 if self._repeat[0] % 2 == 1 else 0
            self._repeat[1] -= 1 if self._repeat[1] % 2 == 1 else 0
        elif not self.repeat.find('单周') == -1:
            self._repeat.append(2)
            self._repeat[0] += 1 if self._repeat[0] % 2 == 0 else 0
            self._repeat[1] -= 1 if self._repeat[1] % 2 == 0 else 0
        else:
            if len(self._repeat) == 1:
                pass
            elif len(self._repeat) == 2:
                self._repeat.append(1)
            else:
                print('repeat error')

    def getRepeat(self):
        rrule = {'FREQ':'WEEKLY', 'INTERVAL':'', 'COUNT':''}
        #从几周到几周，间隔几周

        if len(self._repeat) >= 3:
            rrule['INTERVAL'] = str(self._repeat[-1])
            count = (self._repeat[1] - self._repeat[0]) / self._repeat[-1] + 1
            rrule['COUNT'] = str(int(count))
        else:
            rrule = None

        return rrule

    def getStart(self, area='Xueyuanlu'):
        '''
        Return an arrow object
        '''
        week_off = self._repeat[0] - 1
        day_off = self.time[0]
        hour_off = self.time[1]
        if area == 'Xueyuanlu':
            start = session_begin_time.replace( \
            weeks=+week_off,\
            days=+day_off,  \
            hour=xueyuanlu_time[hour_off][0],\
            minute=xueyuanlu_time[hour_off][1])
        elif area == 'Shahe':
            start = session_begin_time.replace(\
            weeks=+week_off,\
            days=+day_off,  \
            hour=shahe_time[hour_off][0],\
            minute=shahe_time[hour_off][1])
        return start.isoformat()

    def getEnd(self, area='Xueyuanlu'):
        week_off = self._repeat[0] - 1
        day_off = self.time[0]
        if len(self.time) == 2:
            hour_off = self.time[1]
        else: hour_off = self.time[-1]

        if area == 'Xueyuanlu':
            stop = session_begin_time.replace(\
            weeks=+week_off,\
            days=+day_off,  \
            hour=xueyuanlu_time[hour_off][0],\
            minute=xueyuanlu_time[hour_off][1],\
            minutes=+50)  #下课时间加50min
        elif area == 'Shahe':
            start = session_begin_time.replace(\
            weeks=+week_off,\
            days=+day_off,  \
            hour=shahe_time[hour_off][0],\
            minute=shahe_time[hour_off][1],\
            minutes=+50 )   #下课时间加50min
        return stop.isoformat()
    
    def getLocation(self):
        return self.class_room

    def getDescription(self):
        pass

    def getTeacher(self):
        return self.teacher
    

path = '/Users/TimFan/Desktop/1.xls'

data = xlrd.open_workbook(path)

table = data.sheets()[0]
title = table.col_values(0)[0]
if title.find('班级') == -1:
    v1 = r'(?P<class_name>.*?)\<\/br\>(?P<teacher>.*?)\[(?P<repeat>[\w\-\－\，\,]+)\](?P<place_and_time>.*\n.*)'
    v3 = r'(?<=节)\<\/br\>'
    re_celisplit = re.compile(v3)
    re_class = re.compile(v1)
else:
    v3 = r'[\n\t|\<\/br\>]+'
    re_celisplit = re.compile(v3)
    v1 = r'(?P<class_name>.*?)◇(?P<teacher>.*?)\[(?P<repeat>[\w\-\－\，\,]+)\]◇?(?P<place_and_time>[\w\，\,\-\－\(,\),\（,\）]+)'
    re_class = re.compile(v1)
    class_schedule = False
class_info_list = []

for i in range(2,9): #对应周一到周日
    celi_info_list = []
    celi_name_list = []
    for j in range(2,8): #对应第一节到第十二节
        raw = table.col_values(i)[j]
        if raw != '':
            temp = re_celisplit.split(raw)
            celi_info = []
            for celi in temp:
                try:
                    a = re_class.match(celi)
                    if a:
                        re_resualt = a.groupdict()
                        celi_info = ClassInfoHandle(i-2,j-2,re_resualt)
                        celi_info_list.append(celi_info)
                        celi_name = re_resualt['class_name']
                        celi_name_list.append(celi_name)
                    else:
                        raise BaseException('Match Failed: match %s with pattern %s',celi,re_class)
                except Exception as e:
                    logging.exception(e)
                    

    
    '''
    合并连续上的课，todo
    s = [{},]
    s_count = 0
    celi_info_list.sort(key = lambda x:x[1])
    for i in range(len(celi_info_list)-1):
        if i in s[s_count]:
            continue
        else:
            for j in range(i+1,len(celi_info_list)):
                if celi_name_list[i] = celi_name_list[j]:
                    s[s_count].add(i)
                    s[s_count].add(j)
            s_count +=1
            s.append(set())
    for k in s:
        if not len(k)==0:
            k = list(k)
            
   '''
    class_info_list.extend(celi_info_list)
    
            
        



calendar = ics.Calendar()
eventlist = []

for celi in class_info_list:
    
    e = ics.Event()
    rpt = celi.getRepeat()
    e.rrule = rpt if rpt else None
    e.begin = celi.getStart()
    e.end = celi.getEnd()
    e.name = celi.class_name
    e.location = celi.getLocation()
    e.description = celi.getTeacher()
    calendar.events.append(e)


    

output_path = path[0:path.find('.')] + '.ics'

with open(output_path,'a',encoding="utf-8") as f:
    f.writelines(calendar)
print(title)
print('done')