#!/usr/bin/env python

import os
import sys
import itertools
import csv
from string import Template

rowNum = int(sys.argv[1])
print(rowNum)
volunteerPath = "C:/complete/path/to/volunteer_assignments.csv"
jobPath = "C:/complete/path/to/job_descriptions.csv"
templatePath = "C:/complete/path/to/email_template.txt"

with open(jobPath, mode='r') as f:
	reader = csv.reader(f)
	jobLocations = {rows[0]:rows[1] for rows in reader}

with open(volunteerPath, 'r') as f:
	personInfo = next(itertools.islice(csv.reader(f), rowNum, None))

with open(templatePath, 'r') as f:
	template = f.read()

fullname = personInfo[1]
email = personInfo[2]

role8 = personInfo[4]
role9 = personInfo[5]
role10 = personInfo[6]
role11 = personInfo[7]
role12 = personInfo[8]

if role8 == 'FALSE':
	role8 = 'Not Volunteering'
if role9 == 'FALSE':
	role9 = 'Not Volunteering'
if role10 == 'FALSE':
	role10 = 'Not Volunteering'
if role11 == 'FALSE':
	role11 = 'Not Volunteering'
if role12 == 'FALSE':
	role12 = 'Not Volunteering'

location8 = jobLocations[role8]
location9 = jobLocations[role9]
location10 = jobLocations[role10]
location11 = jobLocations[role11]
location12 = jobLocations[role12]

indSchedule = {"fullname": fullname, 
	"role8": role8, 
	"role9": role9,
	"role10": role10,
	"role11": role11,
	"role12": role12,
	"location8": location8,
	"location9": location9,
	"location10": location10,
	"location11": location11,
	"location12": location12
}

newEmail = template.format(**indSchedule)
emailFileName = "./IndividualEmails/" + email + ".txt"
with open(emailFileName, 'w') as f:
	f.write(newEmail)

