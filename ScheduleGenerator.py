#!/usr/bin/python

import sys
import xlsxwriter as xl
import math
import os

courses = []
days = ["Monday","Tuesday","Wednesday","Thursday","Friday"]
times = ["800","830","900","930","1000","1030","1100","1130","1200","1230","1300","1330","1400","1430","1500","1530","1600","1630","1700","1730","1800","1830","1900","1930","2000","2030","2100","2130","2200"]

def RangesOverlap(start1,end1,start2,end2):
	if (start1 > end1 or start2 > end2):
		print("ERROR: VALUES OUT OF ORDER: " + str(start1) + " > " + str(end1) + " OR " + str(start2) + " > " + str(end2))
		raise ValueError
	return ((int(start1) <= int(end2)) and (int(start2) <= int(end1)))

class Section(object):
	def __init__(self,course,sectionID,uniqueID,lecDays,lecStart,lecEnd,labDays,labStart,labEnd,semDays,semStart,semEnd):
		self.course = course
		self.sectionID = sectionID
		self.uniqueID = uniqueID
		self.lecDays = lecDays
		self.lecStart = int(lecStart)
		self.lecEnd = int(lecEnd)
		self.labDays = labDays
		self.labStart = int(labStart)
		self.labEnd = int(labEnd)
		self.semDays = semDays
		self.semStart = int(semStart)
		self.semEnd = int(semEnd)

	def CheckPossibleLectureOverlap(self,otherSection):
		commonLectureDays = set(self.lecDays).intersection(otherSection.lecDays)
		return (len(commonLectureDays) > 0)
	
	def CheckPossibleLabOverlap(self,otherSection):
		commonLabDays = set(self.labDays).intersection(otherSection.labDays)
		return (len(commonLabDays) > 0)

	def CheckPossibleSemOverlap(self,otherSection):
		commonSemDays = set(self.semDays).intersection(otherSection.semDays)
		return (len(commonSemDays) > 0)

	def CheckPossibleLecLabSemOverlap(self,otherSection):
		commonDays = set(self.lecDays + self.labDays + self.semDays).intersection(otherSection.lecDays + otherSection.labDays + otherSection.semDays)
		return (len(commonDays) > 0)

	def CheckPossibleOverlap(self,otherSection):
		f = open("debug.txt","a+")
		print(self.CreateSectionName() + " " + otherSection.CreateSectionName(),file=f)
		lecOverlap = self.CheckPossibleLectureOverlap(otherSection)
		labOverlap = self.CheckPossibleLabOverlap(otherSection)
		semOverlap = self.CheckPossibleSemOverlap(otherSection)
		anyOverlap = self.CheckPossibleLecLabSemOverlap(otherSection)
		print("\tLec " + str(lecOverlap) + "\n\tLab " + str(labOverlap) + "\n\tSem " + str(semOverlap) + "\n\tAny " + str(anyOverlap),file=f)
		if (anyOverlap or lecOverlap or labOverlap or semOverlap):
			print(str(self.lecStart) + " " + str(self.lecEnd) + " " + str(otherSection.lecStart) + " " + str(otherSection.lecEnd),file=f)
			print(str(RangesOverlap(self.lecStart,self.lecEnd,otherSection.lecStart,otherSection.lecEnd)),file=f)
			if (lecOverlap and RangesOverlap(self.lecStart,self.lecEnd,otherSection.lecStart,otherSection.lecEnd)):
				print("\t\tLecSp False -- Overlap",file=f)
				f.close()
				return False
			else:
				print("\t\tLecSp True -- No overlap",file=f)
				if (labOverlap and RangesOverlap(self.labStart,self.labEnd,otherSection.labStart,otherSection.labEnd)):
					print("\t\tLabSp False -- Overlap",file=f)
					f.close()
					return False
				else:
					print("\t\tLabSp True -- No overlap",file=f)
					if (semOverlap and RangesOverlap(self.semStart,self.semEnd,otherSection.semStart,otherSection.semEnd)):
						print("\t\tSemSp False -- Overlap",file=f)
						f.close()
						return False
					else:
						print("\t\tSemSp True -- No overlap",file=f)
						selfLecOtherLabOverlap = True
						if ((self.lecStart != 0 or self.lecEnd != 0) and (otherSection.labStart != 0 or otherSection.labEnd != 0) and (len(set(self.lecDays).intersection(otherSection.labDays)) > 0)):
							#print(str(self.lecStart) + " " + str(self.lecEnd) + " " + str(otherSection.labStart) + " " + str(otherSection.labEnd))
							#print(str(RangesOverlap(self.lecStart,self.lecEnd,otherSection.labStart,otherSection.labEnd)))
							selfLecOtherLabOverlap = RangesOverlap(self.lecStart,self.lecEnd,otherSection.labStart,otherSection.labEnd)
						else:
							selfLecOtherLabOverlap = False
						selfLecOtherSemOverlap = True
						if ((self.lecStart != 0 or self.lecEnd != 0) and (otherSection.semStart != 0 or otherSection.semEnd != 0) and (len(set(self.lecDays).intersection(otherSection.semDays)) > 0)):
							selfLecOtherSemOverlap = RangesOverlap(self.lecStart,self.lecEnd,otherSection.semStart,otherSection.semEnd)
						else:
							selfLecOtherSemOverlap = False
						print("\t\t\t" + str(selfLecOtherLabOverlap) + " " + str(selfLecOtherSemOverlap),file=f)
						
						selfLabOtherLecOverlap = True
						if ((self.labStart != 0 or self.labEnd != 0) and (otherSection.lecStart != 0 or otherSection.lecEnd != 0) and (len(set(self.labDays).intersection(otherSection.lecDays)) > 0)):
							selfLabOtherLecOverlap = RangesOverlap(self.labStart,self.labEnd,otherSection.lecStart,otherSection.lecEnd)
						else:
							selfLabOtherLecOverlap = False
						selfLabOtherSemOverlap = True
						if ((self.labStart != 0 or self.labEnd != 0) and (otherSection.semStart != 0 or otherSection.semEnd != 0) and (len(set(self.labDays).intersection(otherSection.semDays)) > 0)):
							selfLabOtherSemOverlap = RangesOverlap(self.labStart,self.labEnd,otherSection.semStart,otherSection.semEnd)
						else:
							selfLabOtherSemOverlap = False
						print("\t\t\t" + str(selfLabOtherLecOverlap) + " " + str(selfLabOtherSemOverlap),file=f)

						selfSemOtherLecOverlap = True
						if ((self.semStart != 0 or self.semEnd != 0) and (otherSection.lecStart != 0 or otherSection.lecEnd != 0) and (len(set(self.semDays).intersection(otherSection.lecDays)) > 0)):
							selfSemOtherLecOverlap = RangesOverlap(self.semStart,self.semEnd,otherSection.lecStart,otherSection.lecEnd)
						else:
							selfSemOtherLecOverlap = False
						selfSemOtherLabOverlap = True
						if ((self.semStart != 0 or self.semEnd != 0) and (otherSection.labStart != 0 or otherSection.labEnd != 0) and (len(set(self.semDays).intersection(otherSection.labDays)) > 0)):
							selfSemOtherLabOverlap = RangesOverlap(self.semStart,self.semEnd,otherSection.labStart,otherSection.labEnd)
						else:
							selfSemOtherLabOverlap = False
						print("\t\t\t" + str(selfSemOtherLecOverlap) + " " + str(selfSemOtherLabOverlap),file=f)

						'''
						if (anyOverlap and selfLecOtherLabOverlap and selfLecOtherSemOverlap and selfLabOtherLecOverlap and selfLabOtherSemOverlap and selfSemOtherLecOverlap and selfSemOtherLabOverlap):
							#print("Overlap\n")
							return False
						else:
							#print("No overlap\n")
							return True
						'''
						print(str(selfLecOtherLabOverlap) + " " + str(selfLecOtherSemOverlap) + " " + str(selfLabOtherLecOverlap) + " " + str(selfLabOtherSemOverlap) + " " + str(selfSemOtherLecOverlap) + " " + str(selfSemOtherLabOverlap),file=f)
						if (selfLecOtherLabOverlap or selfLecOtherSemOverlap or selfLabOtherLecOverlap or selfLabOtherSemOverlap or selfSemOtherLecOverlap or selfSemOtherLabOverlap):
							print("Overlap\n",file=f)
							f.close()
							return False
						else:
							print("No overlap\n",file=f)
							f.close()
							return True
		
		#print("Overlap\n")
		#return True
		
		print("No overlap\n",file=f)
		f.close()
		return True

	def CreateSection(course,sL):
		sectionID = sL[2]
		uniqueID = sL[3]

		lecStartTime = sL[5]
		lecEndTime = sL[6]

		numLecDays = int(sL[7])
		lecDays = []

		for i in range(1,numLecDays+1):
			lecDays.append(sL[7+i])

		numLabDays = int(sL[11+numLecDays])
		labStartTime = 0
		labEndTime = 0
		labDays = []
		if (numLabDays > 0):
			labStartTime = sL[9+numLecDays]
			labEndTime = sL[10+numLecDays]

			for i in range(1,numLabDays+1):
				labDays.append(sL[11+numLecDays+i])
		
		numSemDays = int(sL[15+numLecDays+numLabDays])
		semStartTime = 0
		semEndTime = 0
		semDays = []
		if (numSemDays > 0):
			semStartTime = sL[13+numLecDays+numLabDays]
			semEndTime = sL[14+numLecDays+numLabDays]

			for i in range(1,numSemDays+1):
				semDays.append(sL[15+numLecDays+numLabDays+i])

		newSection = Section(course,sectionID,uniqueID,lecDays,lecStartTime,lecEndTime,labDays,labStartTime,labEndTime,semDays,semStartTime,semEndTime)
		return newSection

	def CreateSectionName(self):
		return str(str(self.course.name.upper()) + " " + str(self.sectionID) + " (" + str(self.uniqueID) + ")")

	def FindCompatibleSections(self,sections):
		compatibleSections = []
		for oSection in sections:
			if (self != oSection and self.course != oSection.course):
				if (self.CheckPossibleOverlap(oSection)):
					compatibleSections.append(oSection)
		return compatibleSections

class Course(object):
	def __init__(self,courseName):
		self.name = courseName
		self.sections = []

	def AddSection(self,section):
		self.sections.append(section)

def ParseCourses(fileName):
	file = open(fileName,'r')
	i = 0
	for line in file:
		line = line.split('\n')[0]
		if (i != 0):
			splitLine = line.split(',')
			courseName = splitLine[0] + " " + splitLine[1]
			foundCourse = False
			for course in courses:
				if (course.name == courseName):
					foundCourse = True
					newCourse.AddSection(Section.CreateSection(newCourse,splitLine))
					break
			if (foundCourse == False):
				newCourse = Course(courseName)
				newCourse.AddSection(Section.CreateSection(newCourse,splitLine))
				courses.append(newCourse)
		i += 1
	file.close()
	if (len(courses) < 2):
		print("\nScheduling " + str(len(courses)) + " course"+("" if len(courses) == 1 else "s") + " is not supported, must have more than 1. Please modify your csv file.")
		sys.exit()

def GenerateSingleSchedules(schedule,schedules,compatibleSections,index):
	for section in compatibleSections:
		schedule.append(section)
		if (index == len(courses)-1):
			schedules.append(schedule.copy())
			del schedule[index]
			continue
		innerCompatibleSections = section.FindCompatibleSections(compatibleSections)
		if (len(innerCompatibleSections) > 0):
			GenerateSingleSchedules(schedule,schedules,innerCompatibleSections,index+1)
			del schedule[len(schedule)-1]
		else:
			del schedule[index]
			continue
	return

def CreateScheduleOutput(schedule):
	output = ""
	for i in range(len(schedule)):
		output += schedule[i].uniqueID
		if (i != len(schedule)-1):
			output += ","
	return output

def GenerateSchedules():
	schedules = []
	progressCount = 0
	progressTotal = 0
	for course in courses:
		for section in course.sections:
			progressTotal += 1
	for course in courses:
		for section in course.sections:
			compatibleSections = []
			tempSchedule = [section]
			for otherCourse in courses:
				if (course != otherCourse):
					for otherSection in otherCourse.sections:
						if (section.CheckPossibleOverlap(otherSection)):
							compatibleSections.append(otherSection)
			GenerateSingleSchedules(tempSchedule,schedules,compatibleSections,1)

			progressCount += 1
			print("\r" + str(int(round((progressCount/progressTotal)*100,0))) + "%",end="\r")
	for schedule in schedules:
		schedule.sort(key=lambda section: section.course.name)
	schedules = list(set(tuple(schedule) for schedule in schedules))
	if (len(schedules) <= 0):
		print("\nThere are no possible schedules that work with the courses provided.")
		sys.exit()
	return schedules

def Convert24Hto12H(time):
	time = int(time)
	if (time >= 1300):
		time -= 1200
	time = list(str(time))
	if (len(time) == 3):
		time.insert(1,':')
	elif (len(time) == 4):
		time.insert(2,':')
	return ''.join(time)

def RangesCompletelyOverlap(start1,end1,start2,end2):
	return (int(end1)-int(end2) <= 0 and int(start2)-int(start1) <= 0)

def CheckWithinTimeRange(section,startTime,endTime):
	if (RangesCompletelyOverlap(section.lecStart,section.lecEnd,startTime,endTime)):
		if ((section.labStart == 0 or section.labEnd == 0) or RangesCompletelyOverlap(section.labStart,section.labEnd,startTime,endTime)):
			if ((section.semStart == 0 or section.semEnd == 0) or RangesCompletelyOverlap(section.semStart,section.semEnd,startTime,endTime)):
				return True
			else:
				return False
		else:
			return False
	else:
		return False

def CompareTimes(section,startTimesPerDay,endTimesPerDay,index,day):
	if (day == days[index]):
		if (day in section.lecDays):
			if (section.lecStart != 0 and section.lecStart < startTimesPerDay[index]):
				startTimesPerDay[index] = section.lecStart
			if (section.lecEnd != 0 and section.lecEnd > endTimesPerDay[index]):
				endTimesPerDay[index] = section.lecEnd
		if (day in section.labDays):
			if (section.labStart != 0 and section.labStart < startTimesPerDay[index]):
				startTimesPerDay[index] = section.labStart
			if (section.labEnd != 0 and section.labEnd > endTimesPerDay[index]):
				endTimesPerDay[index] = section.labEnd
		if (day in section.semDays):
			if (section.semStart != 0 and section.semStart < startTimesPerDay[index]):
				startTimesPerDay[index] = section.semStart
			if (section.semEnd != 0 and section.semEnd > endTimesPerDay[index]):
				endTimesPerDay[index] = section.semEnd

def FindClosestValueInIntegerArray(value,array):
	if (len(array) > 0):
		closestValue = array[0]
		for arrayValue in array:
			if (abs(int(value)-int(arrayValue)) < abs(int(value)-int(closestValue))):
				closestValue = arrayValue
		return int(closestValue)
	return 0

def ExcelOutput(section,day,formats,worksheet,start,end,text):
	row = times.index(str(start))+1
	col = days.index(str(day))+1
	worksheet.set_column(row,col,20)

	borderThicknessIndex = 2 # 0, 1, 2, 5

	format1 = formats[0]
	format1.set_left(borderThicknessIndex)
	format1.set_right(borderThicknessIndex)

	format2 = formats[1]
	format2.set_left(borderThicknessIndex)
	format2.set_right(borderThicknessIndex)
	format2.set_top(borderThicknessIndex)

	format3 = formats[2]
	format3.set_left(borderThicknessIndex)
	format3.set_right(borderThicknessIndex)
	format3.set_bottom(borderThicknessIndex)

	worksheet.write(row,col,section.course.name.upper() + " " + text,format2)
	nextTime = times[row-1]
	k = 0
	while ((int(nextTime)-10) < int(end) and row <= len(times)-1):
		row += 1
		nextTime = times[row]
		if (k == 0):
			worksheet.write(row,col,section.sectionID + " (" + section.uniqueID + ")",format1)
		elif (k == times.index(str(FindClosestValueInIntegerArray(end,times)))-times.index(str(FindClosestValueInIntegerArray(start,times)))-2):
			worksheet.write(row,col,"",format3)
		else:
			worksheet.write(row,col,"",format1)
		k += 1

def ChooseSchedule(schedules):
	startTime = "0700"
	endTime = "2400"
	validSchedules = []
	retry = "Yes"
	while (retry == "Yes"):
		validSchedules.clear()
		retry = "No"
		while True:
			try:
				startTime = input("\nStart after (24-hour) (currently: " + str(startTime) + "): ")
				int(startTime)
				if (int(startTime) > 2400 or (int(startTime) < 0)):
					raise ValueError
			except:
				print("\n\"" + str(startTime) + "\" is not a valid time. Try again.")
				startTime = "0700"
				continue
			else:
				break
		while True:
			try:
				endTime = input("End before (24-hour) (currently: " + str(endTime) + "): ")
				int(endTime)
			except:
				print("\n\"" + str(endTime) + "\" is not a 24-hour valid time. Try again.")
				endTime = "2400"
				continue
			else:
				break
		numSchedulesWork = 0
		noSectionsWork = True
		for schedule in schedules:
			sectionsWithinRange = True
			for section in schedule:
				if (not CheckWithinTimeRange(section,int(startTime),int(endTime))):
					sectionsWithinRange = False
			if (sectionsWithinRange):
				validSchedules.append(schedule)
				numSchedulesWork += 1
				noSectionsWork = False
				startTimesPerDay = [9999,9999,9999,9999,9999]
				endTimesPerDay = [0,0,0,0,0]
				for section in schedule:
					combinedDays = section.lecDays + section.labDays + section.semDays
					for i in range(len(combinedDays)):
						CompareTimes(section,startTimesPerDay,endTimesPerDay,days.index(combinedDays[i]),combinedDays[i])
				if (numSchedulesWork <= 10):
					scheduleString = ""
					for i in range(len(days)):
						scheduleString += ("\t" if (i == 0) else " ") + str(days[i][:2]) + ": " + str(Convert24Hto12H(startTimesPerDay[i])) + "-" + str(Convert24Hto12H(endTimesPerDay[i]) + (" | " if (i != len(days)-1) else ""))
					print("\nSchedule " + str(numSchedulesWork) + "\t" + scheduleString)
					for section in schedule:
						lecString = "\t" + section.CreateSectionName()
						formattedLecDays = ""
						formattedLabDays = ""
						formattedSemDays = ""
						for i in range(len(section.lecDays)):
							formattedLecDays += (str(''.join(list(section.lecDays[i])[:2])))
							if (i < len(section.lecDays)-1):
								formattedLecDays += (", ")
						lecString += "\tLectures (" + formattedLecDays + "):\t" + str(Convert24Hto12H(section.lecStart)) + "-" + str(Convert24Hto12H(section.lecEnd))
						print(lecString)
						labString = ""
						for i in range(len(section.labDays)):
							formattedLabDays += (str(''.join(list(section.labDays[i])[:2])))
							if (i < len(section.labDays)-1):
								formattedLabDays += (", ")
						if (len(section.labDays) > 0):
							labString += "\t\t\t\tLabs     (" + formattedLabDays + "):\t\t" + str(Convert24Hto12H(section.labStart)) + "-" + str(Convert24Hto12H(section.labEnd)) + "\t"
							print(labString)
						semString = ""
						for i in range(len(section.semDays)):
							formattedSemDays += (str(''.join(list(section.semDays[i])[:2])))
							if (i < len(section.semDays)-1):
								formattedSemDays += (", ")
						if (len(section.semDays) > 0):
							semString += "\t\t\t\tSeminars (" + formattedSemDays + "):\t\t" + str(Convert24Hto12H(section.semStart)) + "-" + str(Convert24Hto12H(section.semEnd)) + "\t"
							print(semString)
		if (noSectionsWork):
			retry = input("\nThere are no possible schedules with the start and end times provided. Retry? (Yes/No): ")
		else:
			print("\nThere are " + str(numSchedulesWork) + " schedules that fit within those times.")
			retry = input("\nWould you like to change your search? (Yes/No): ")
	while (True):
		excelCreate = input("Would you like to see an excel visualisation of a schedule? (Yes/No): ")
		if (excelCreate == "Yes"):
			scheduleInput = ""
			while True:
				try:
					scheduleInput = input("\nWhich schedule would you like to see an excel visualisation of? (Schedule Number): ")
					int(scheduleInput)
					if (int(scheduleInput)-1 >= len(validSchedules) or int(scheduleInput)-1 < 0):
						raise IndexError
				except ValueError:
					print("Invalid schedule number: \"" + scheduleInput + "\" is not a number.")
					continue
				except IndexError:
					print("Invalid schedule number: " + scheduleInput + " is not a valid schedule choice. There are " + str(len(validSchedules)) + " schedules.")
				else:
					break
			workbookFileName = "Schedule"+scheduleInput+".xlsx"
			scheduleInput = str(int(scheduleInput)-1)
			workbook = xl.Workbook(workbookFileName)
			worksheet = workbook.add_worksheet()

			cellColours = ["#D35400","#2C3E50","#8E44AD","#2980B9","#27AE60","#C0392B","#16A085"]
			titleColours = ["#ECF0F1","#DAE0E5"]

			for i in range(len(days)):
				format = workbook.add_format()
				format.set_bg_color(titleColours[1])
				worksheet.write(0,i+1,days[i],format)
			for i in range(len(times)):
				format = workbook.add_format()
				colourIndex = 0
				if (i == 0 or i % 2 == 0):
					colourIndex = 0
				else:
					colourIndex = 1
				format.set_bg_color(titleColours[colourIndex])
				worksheet.write(i+1,0,Convert24Hto12H(times[i]),format)
			for i in range(len(validSchedules[int(scheduleInput)])):
				section = validSchedules[int(scheduleInput)][i]

				format = workbook.add_format()
				format.set_bg_color(cellColours[i])
				format.set_font_color("#FFFFFF")

				format2 = workbook.add_format()
				format2.set_bg_color(cellColours[i])
				format2.set_font_color("#FFFFFF")

				format3 = workbook.add_format()
				format3.set_bg_color(cellColours[i])
				format3.set_font_color("#FFFFFF")

				print(section.CreateSectionName())
				for day in section.lecDays:
					ExcelOutput(section,day,[format,format2,format3],worksheet,section.lecStart,section.lecEnd,"Lecture")
				for day in section.labDays:
					ExcelOutput(section,day,[format,format2,format3],worksheet,section.labStart,section.labEnd,"Lab")
				for day in section.semDays:
					ExcelOutput(section,day,[format,format2,format3],worksheet,section.semStart,section.semEnd,"Seminar")
			while True:
				try:
					workbook.close()
				except PermissionError:
					tryAgain = input("The file is already open or you have incorrect permissions to edit the file. Close the file and try again. (Press Enter to Continue)")
					continue
				else:
					break
			import os
			if os.name == 'nt':
				while True:
					try:
						os.startfile(workbookFileName)
						print("Opening " + workbookFileName)
					except PermissionError:
						tryAgain = input("The file is already open or you have incorrect permissions to open the file. Close the file and try again. (Press Enter to Continue)")
						continue
					else:
						break
		else:
			break
		
def main():
	if (len(sys.argv) != 2):
		print("Incorrect argument format.\nCorrect format: python " + sys.argv[0] + " fileName.csv")
		sys.exit()
	
	print("Parsing course file...")
	ParseCourses(sys.argv[1])
	print("Generating schedule combinations...")
	schedules = GenerateSchedules()
	print("There are " + str(len(schedules)) + " different schedules combinations.")
	ChooseSchedule(schedules)

main()