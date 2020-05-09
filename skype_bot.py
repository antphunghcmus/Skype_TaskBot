import sys, io, string, datetime, re, time, asyncio
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from skpy import Skype, SkypeMsg

sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding = 'utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding = 'utf-8')

async def main():
	# use creds to create a client to interact with the Google Drive API
	scope = ['https://spreadsheets.google.com/feeds',
			 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('./client_secret.json', scope)
	client = gspread.authorize(creds)

	# Open login file
	login_file = open('./login.txt')
	lines = login_file.readlines()
	username = lines[0].rstrip()
	password = lines[1].rstrip()

	# Open skype group file
	skypeGroup_file = open('./config.txt')
	skypeGroup_arr = skypeGroup_file.readlines()

	# connect to Skype
	sk = Skype(username,password)
	mysuser = sk.user # you
	ch=sk.chats

	#Call back to handle message in all groups
	while True:
		for group_link in skypeGroup_arr[1:]:
			await async_group_message(client, ch, group_link)
		time.sleep(5)

def update_task(userID, task, worksheet):
	# Find user
	user = worksheet.find(userID)
	column = int(user.col) + 1
	old_task = worksheet.cell(4, column).value #Get old task
	if old_task == '':
		new_task = "'+ " + task
	else:
		new_task = old_task + chr(10) + "+ " + task
	print(userID)
	print(new_task)
	find_assign_task = worksheet.update_cell(4, column, new_task) # Update using cell

async def async_group_message(client, ch, group_link):
	groupid= ch.urlToIds(group_link)
	ch2= ch.chat(groupid["id"])
	while True:
		now =datetime.datetime.utcnow()
		fiveminute= now - datetime.timedelta(minutes=5) # time to get task
		Mess= ch2.getMsgs() # Get 8 mess
		for j in Mess:
			if j.time >= fiveminute:    # in last 5 minute
				if j.content.find('#') == 0:  #task
					print(j.content)
					tasks = fillter_task(j)
					#Find a workbook and open sheet by name
					dt = datetime.datetime.today()
					sheet_name = str(dt.day) + '/' + str(dt.month)
					worksheet = client.open("Tempo 5000").worksheet(sheet_name)
					for task in tasks:
						update_task(task['userID'], task['task_content'], worksheet)
						alert_skype(j.user.name, task['userID'], task['task_content'], ch, True)
						print('Sent task to: ', task['userID'])
				else:
					continue
		break

def fillter_task(j):
	task_total = []
	task_key='# <at id="8:'
	hashtag_key = '<at id="8:'
	content = j.content
	while True:
		if content.find(task_key) != -1:
			start = content.find(task_key) + len(task_key)
			end = content.find('">')
			userID = content[start:end] # Find user id
			content = content[end + 2:]	# Update content
			start = content.find('</at>') + 5 # Find task content of UserID above
			print(content)
			print(content.find(hashtag_key))
			if content.find(hashtag_key) != -1:
				end = content.find(hashtag_key)
				task_content = content[start:end] + ' ' + str(j.time)# Find task
				new_task = {
					"userID" : userID,
					"task_content" : task_content
				}
				task_total.append(new_task)
				content = content[end:]
				task_key = hashtag_key #Update key
			else:
				task_content = content[start:] + ' ' + str(j.time)# Find task
				new_task = {
					"userID" : userID,
					"task_content" : task_content
				}
				task_total.append(new_task)
				content=''
		else:
			return task_total

def alert_skype(leaderName,userID,Task, ch, status, user=None, time=None, content=None):
	idSkype='8:'+userID
	ch1= ch.chat(idSkype) # connect to chat using id
	now =datetime.datetime.now()
	if status == True:
		alert = 'You have a new Task from ' + SkypeMsg.bold(str(leaderName)) + chr(10) + SkypeMsg.bold(str(Task))
		ch1.sendMsg(alert, rich=True)
	else:
		alert = SkypeMsg.quote(user, ch1, time1,content1)  + chr(10) + 'syntax error'
		ch1.sendMsg(alert, rich=True)

if __name__ == '__main__':
	asyncio.run(main())