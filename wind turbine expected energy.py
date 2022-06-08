import pandas as pd
from datetime import datetime, timedelta
import os
import time
file = "C:/Users/username/Desktop/Boook2.xlsx"
# 'windspeed' , 'Feuil2'

df = pd.read_excel(file , sheet_name='HistoryAlarmLog')
dates = df[['TimeOn' , 'TimeOff']]
df2 = pd.read_excel(file , sheet_name='windspeed')
controls =  df2[['TimeStamp','enery en second']]
index = df.index
number_of_rows_df = len(index)
s = 0
n = number_of_rows_df
while  s < number_of_rows_df :
	counter = 0
	startdate = datetime.strptime(str(dates.iloc[s,0]), '%Y-%m-%d %H:%M:%S') 
	# print(startdate)
	enddate = datetime.strptime(str(dates.iloc[s,1]), '%Y-%m-%d %H:%M:%S')
	# print(enddate)

	def hour_rounder(t):
		discard = timedelta(minutes=t.minute % 10,
		                             seconds=t.second,
		                             microseconds=t.microsecond)
		t -= discard
		if discard >= timedelta(minutes=10):
		    t += timedelta(minutes=10)
		return t 

	flter = controls[(controls['TimeStamp'] >= hour_rounder(startdate)) & (controls['TimeStamp'] <= hour_rounder(enddate))]
	index = flter.index
	print(flter)
	number_of_rows = len(index)
	# print(flter.iloc[1:-1,:])
	if number_of_rows == 1:
		start = enddate - startdate
		# print(start , 'start date')
		x = time.strptime(str(start).split(',')[0],'%H:%M:%S')
		seconds = timedelta(hours=x.tm_hour,minutes=x.tm_min,seconds=x.tm_sec).total_seconds()
		# print(seconds , 'seconds')
		counter += seconds * flter.iloc[0,1]
		# print(flter)


	if number_of_rows == 2:
		start1 = flter.iloc[1,0] - startdate

		start = str(start1).split(" ")[2]
		x = time.strptime(str(start).split(',')[0],'%H:%M:%S')
		seconds = timedelta(hours=x.tm_hour,minutes=x.tm_min,seconds=x.tm_sec).total_seconds()

		counter += seconds * flter.iloc[0,1]
		end1 = enddate - flter.iloc[-1,0]
		end = str(end1).split(" ")[2]

		x1 = time.strptime(str(end).split(',')[0],'%H:%M:%S')
		seconds1 = timedelta(hours=x1.tm_hour,minutes=x1.tm_min,seconds=x1.tm_sec).total_seconds()

		counter += seconds1 * flter.iloc[1,1]


	if number_of_rows >= 3:
		# counted  start  date  diffent  between  start and  round to 10 min
		# start = startdate - hour_rounder(startdate)
		start1 = flter.iloc[1,0] - startdate
		start = str(start1).split(" ")[2]
		x = time.strptime(str(start).split(',')[0],'%H:%M:%S')
		seconds = timedelta(hours=x.tm_hour,minutes=x.tm_min,seconds=x.tm_sec).total_seconds()
		counter += seconds * flter.iloc[0,1]
		# counted  end  date  diffent  between  end  and  round to 10 min
		end = enddate - hour_rounder(enddate)
		x1 = time.strptime(str(end).split(',')[0],'%H:%M:%S')
		seconds1 = timedelta(hours=x1.tm_hour,minutes=x1.tm_min,seconds=x1.tm_sec).total_seconds()
		counter += seconds1 * flter.iloc[-1,1]
		# print(flter.iloc[-1,1])
		lastrows = flter.iloc[1:-1,:]
		for index, row in lastrows.iterrows():
			# print(row["enery en second"])
			counter+= row["enery en second"]*600
	if number_of_rows ==0:
		pass
	# if  you want  to  append  to  new  sheet  with  results  
	df.loc[s, 'energy expected '] = counter
	print(counter)
	s+=1
df.to_excel(r'C:\\Users\\aswan\\Desktop\\results.xlsx')