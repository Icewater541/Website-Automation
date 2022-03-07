from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path
from plyer import notification
from datetime import date

global niceDate, monthday, dashsearchnames, emails, web, playersnames, midfield, defender, forward

#Initialize the golbal variables
niceDate = date.today()
monthday = date.today().strftime(("%m-%d"))  # create the date formatted as only month and day
dashsearchnames = []
emails = []
playersnames = []
web = webdriver.Chrome()

#Initialize other variables
defender = []
midfield = []
forward = []
centerback = []
outsideback = []
midfield = []
winger = []
striker = []
nicknames = 0

f = open("FILE_NAME", "r")
lines = f.readlines()
f.close()

for x in lines:
    line = x.strip().split('-')

    if line[0][0] != '#' and nicknames == 0:
        dashsearchnames.append(line[0])
        playersnames.append((line[0]))
        emails.append((line[1]))

        if line[2] == "D":
            if line[3] == "o":
                outsideback.append(line[0])
            else:
                centerback.append((line[0]))

        if line[2] == "M":
            midfield.append(line[0])

        if line[2] == "F":
            if line[3] == "s":
                striker.append(line[0])
            else:
                winger.append((line[0]))

    if nicknames == 1:
        for x in range(len(playersnames)):
            if playersnames[x] == line[0]:
                playersnames[x] = line[1]

    if line[0] == '###':
        nicknames = 1

defender = outsideback + centerback
forward = striker + winger

"""
exportPlayer goes through the necessary steps to export the player's pdf
Receives 'name'
name - name of the player
"""
def exportPlayer(name):
    time.sleep(1)
    # try:
    player = web.find_element_by_xpath('/html/body/div[5]/div/div[3]/div[2]/div/div/div/div/div[1]/div[2]/div/div[4]/div/div[2]/div/div[1]/div/div[1]/div[1]/div[3]/input')
    # except:
    #  notification.notify("Error", "Error with finding element. Waiting then try again")
    #  time.sleep(5)
    #  player = web.find_element_by_xpath('/html/body/div[5]/div/div[3]/div[2]/div/div/div/div/div[1]/div[2]/div/div[4]/div/div[2]/div/div[1]/div/div[1]/div[1]/div[3]/input')
    #time.sleep(1)
    player.clear()          #Get rid of current player
    time.sleep(1)
    player.send_keys(name)  #put in player's name to search
    time.sleep(.5)
    #Here, player's name should be the first option shown. We can press arrow down and then enter to select that player
    player.send_keys(Keys.ARROW_DOWN)
    #time.sleep(1)
    player.send_keys(Keys.RETURN)
    dash = clearDash()   #clear the dash name
    time.sleep(.5)
    dash.send_keys(name)    #put player's name into dash
    #time.sleep(2)
    #find the player element to change


    export = web.find_element_by_xpath('//*[@id="dash_editmode"]/div/div[6]/div[1]/button') #find the export button
    time.sleep(.8)
    export.click()      #click export button
    time.sleep(2)       #wait for pdf to export

"""
archiveFiles takes all of he files that have been downloaded into the downloads folder and moves them into a folder
where they can easily be archived
"""
def archiveFiles():
    #loop through all players to move files
    for i in range(len(dashsearchnames)):
        #move the files by renaming them.
        #moves the files from downloads to the archived folder that was created for today
        os.rename('FILE_NAME',
                  'FILE_NAME')

"""
clearDash clears the stuff in the dashboard
"""
def clearDash():
    dash = web.find_element_by_xpath('//*[@id="dash_editmode"]/div/div[1]/div/input')   #find the dashboard
    dash.send_keys(Keys.CONTROL + "a")  #select everything in the dash
    dash.send_keys(Keys.DELETE)         #delete it all
    time.sleep(1)
    return dash

"""
reset adjusts current dash to back the way it was when we first accessed it
Receives 'selection', 'profile', 'position'
selection - the element that needs to be changed
profile - the original name of the dash
position - original player position i.e. Winger, Outside Back, etc.
"""
def reset(selection, profile, position):
    time.sleep(1)
    dash = clearDash()   #Clear the dash
    dash.send_keys(profile) #put the original dash name in
    time.sleep(1)
    selection.send_keys(position)   #set the position back to original
    time.sleep(1)
    #position is only selection, hit arrow down then enter
    selection.send_keys(Keys.ARROW_DOWN)
    time.sleep(1)
    selection.send_keys(Keys.RETURN)
    time.sleep(2)

"""
changeField adjusts player position i.e. Outside Back to Center Back
Receives 'web', 'selection', 'position'
selection - the element that needs to be changed
position - position that we want to change to
"""
def changeField(selction, position):
    #works same as end of reset - put name in and arrow down then enter
    time.sleep(1)
    selection.send_keys(position)
    time.sleep(1)
    selection.send_keys(Keys.ARROW_DOWN)
    time.sleep(1)
    selection.send_keys(Keys.RETURN)

"""
login logs the user into the admin dashboard
This is where username and password is stored
"""
def login():
    username = "USERNAME"
    password = "PASSWORD"
    user = web.find_element_by_xpath('//*[@id="email"]')
    user.send_keys(username)

    pWord = web.find_element_by_xpath('//*[@id="password"]')
    pWord.send_keys(password)

    submit = web.find_element_by_xpath('//*[@id="submit"]')
    submit.click()

"""
Sends an email to another person with the pdf attached
Receives 'email_recipient', 'email_subject', 'email_message', 'attachment_location'
email_recipient - person who will receive the email
email_subject - Subject line of email
email_message - contents of email
attachment_location - directory of where the pdf is stored
"""
def send_email(email_recipient, email_subject, email_message, attachment_location = ''):
    email_sender = 'SENDERE MAIL ADDRESS'  #person sending email

    #set up email to be sent
    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject

    #add the content to the email
    msg.attach(MIMEText(email_message, 'plain'))

    #add attachement if one was specified
    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(attachment_location, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        "attachment; filename= %s" % filename)
        msg.attach(part)

    #attempt to send the email to the recipient
    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.login('EMAIL', 'PASSWORD')
        text = msg.as_string()
        server.sendmail(email_sender, email_recipient, text)
        #print('email sent')
        server.quit()
    except:
        print("SMPT server connection error")
    return True

"""
sendEmails loops through each player and coach and then sends them emails with report attached
"""
def sendEmails():
    #address to be addedd to emails
    address = "SIGNATURE OF EMAIL"
    today = date.today()                    #get the date of today in year, day, moth
    formatteddate = date.today().strftime("%B %d, %Y")  #date that is spelled out i.e March 13, 2021

    #content of the email sent to players - NOTE: \n is the character equivilent of 'Enter'
    message = "MESSAGE TO DISPLAY IN BODY OF EMAIL"

    #loop through every player and send them an email
    for x in range(len(dashsearchnames)):
        sEmail = emails[x] + address    #email address to send to
        #location where the pdf player profile is stored
        location = "FILE LOCATION"
        #Send the email to the player
        send_email(sEmail,
                   '' + str(formatteddate) + ' Daily Player Report',
                   'Hey ' + playersnames[x] + message,
                   location)

"""
createFolder creates a folder with the current date as the name
"""
def createFolder():
    created = "FILENAME"    #location where folder is to be created
    os.mkdir(created)   #Create the folder using the location

"""
FUNCTIONS ARE DONE BEING CREATED, MAIN PROGRAM TO FOLLOW
"""

#get the browser and open the address of the dashboard
web.get("WEBSITE ADDRESS OF STATISTICS PROVIDER")
login() #Log into email

#Look for dash selection box and then get all possible options and store them in all_options
element = web.find_element_by_xpath("/html/body/div[5]/div/div[3]/div[1]/div[2]/select")
all_options = element.find_elements_by_tag_name("option")

#loop through each possible dashboard
for option in all_options:
    #This takes care of the dash for mifielders
    if option.text.strip() == "Midfield Profile":
        option.click() #select profile
        #loop through each midfielder who needs a profile
        for i in range(len(midfield)):
            exportPlayer(midfield[i])  #Export the player's profile
        #Loop has finsihed, all midfielder profiles downloaded
        dash = clearDash()   #Clear the dash
        dash.send_keys("Midfield Profile")  #reset dash to original name


    #This takes care of the dash for forwards
    if option.text.strip() == "Forward Profile":
        option.click()  #select profile
        #loop through each forward who needs a profile
        for i in range(len(forward)):
            if i < len(striker):   #The first x players who are strikers. Just regualr export
                exportPlayer(forward[i])
            else:   #Onto wingers now, strikers are done
                if i == len(striker):  #change "striker" field to "winger" field
                    selection = web.find_element_by_xpath('//*[@id="pLPuZrd9MN"]/div[2]/div/div[1]/div[1]/div[3]/input')
                    changeField(selection, "Winger")
                exportPlayer(forward[i])   #export player profile
        time.sleep(1)
        dash = clearDash()   #Clear dashname and reset dashboard to original state
        reset(selection, "Forward Profile", "Striker")

    #This takes care of the dash for defenders
    if option.text.strip() == "Defender Profile":
        # select profile
        option.click()
        # loop through each forward who needs a profile
        for i in range(len(defender)):
            if i < len(outsideback):   #The first x players who are outside backs. Just regualr export
                exportPlayer(defender[i])
            else:   #Onto Center Backs now, Outside Backs are done
                if i == len(outsideback):  #change "Outside Back" field to "Center Back" field
                    selection = web.find_element_by_xpath('//*[@id="oCENgMdP7j"]/div[2]/div/div[1]/div[1]/div[3]/input')
                    changeField(selection, "Center Back")
                exportPlayer(defender[i])  #Export player profile
        time.sleep(1)
        #Clear dash name and then reset the dasboard to the original state
        dash = clearDash()
        reset(selection, "Defender Profile", "Outside Back")

createFolder()      #Create the new folder where pdfs will go
archiveFiles()      #Move the downloaded files into the new folder
time.sleep(2)
web.close()         #Close the browswer that we were using
web.quit()          #Get rid of the instance of the webdriver
sendEmails()        #Send out the emails to all of teh players and coaches

#Let us know that the emails have been sent out and that the files are stored
notification.notify("Emails", "Done sending emails and files archived")