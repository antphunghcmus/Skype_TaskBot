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
	for group_link in skypeGroup_arr[1:]:
		await asyncio.gather(async_group_message(client, ch, group_link))

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
					start = j.content.find('id="8:') + 6
					end = j.content.find('">')
					userID = j.content[start:end] # Find user id
					start = j.content.find('</at>') + 6
					task = j.content[start:] + ' ' + str(j.time)# Find task
					print(userID)
					print(task)
					# Find a workbook and open sheet by name
					dt = datetime.datetime.today()
					sheet_name = str(dt.day) + '/' + str(dt.month)
					worksheet = client.open("Tempo 5000").worksheet(sheet_name)
					update_task(userID, task, worksheet)
					alert_skype(j.user.name, userID, task, ch)
					print('Sent task to: ',userID)
				else:
					continue
		time.sleep(5)

def alert_skype(LeaderName,UserID,Task, ch):
	idSkype='8:'+UserID
	ch1= ch.chat(idSkype) # connect to chat using id
	now =datetime.datetime.now()
	alert = 'You have a new Task from ' + SkypeMsg.bold(str(LeaderName)) + chr(10) + SkypeMsg.bold(str(Task))
	ch1.sendMsg(alert, rich=True)

if __name__ == '__main__':
	asyncio.run(main())