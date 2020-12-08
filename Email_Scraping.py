import imaplib
import email
import getpass
import pandas as pd
from datetime import datetime


"""
Desired Email address is entered as a string
Then login to the email server
"""
username = "YourEmail@gmail.com"
password = ("YourPassword")
mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(username, password)

print(mail.list())
mail.select("inbox")

result, numbers = mail.uid('search', None, "ALL")
uids = numbers[0].split()
uids = [id.decode("utf-8") for id in uids ]
uids = uids[-1:-3000:-1]
result, messages = mail.uid('fetch', ','.join(uids), '(BODY[HEADER.FIELDS (SUBJECT FROM DATE)])')

date_list = []
from_list = []
from_list1 = []
date_list1 = []


subject_text = []
for i, message in messages[::2]:
    msg = email.message_from_bytes(message)
    decode = email.header.decode_header(msg['Subject'])[0]
    if isinstance(decode[0],bytes):
        decoded = decode[0].decode(errors='ignore')
        subject_text.append(decoded)
    else:
        subject_text.append(decode[0])
    date_list.append(msg.get('date'))
    fromlist = msg.get('From')
    fromlist = fromlist.split("<")[0].replace('"', '')
    from_list1.append(fromlist)
date_list = pd.to_datetime(date_list)
date_list1 = []
for item in date_list:
    date_list1.append(item.isoformat(' ')[:-6])
print(len(subject_text))
print(len(from_list1))
print(len(date_list1))
df = pd.DataFrame(data={'Date':date_list1, 'Sender':from_list1, 'Subject':subject_text})
print(df.head())



df.to_csv(r'C:\Users\User\Desktop\email_inbox.csv',index=False)

df.describe()

"""Using Datetime to create new values
SinceMid is the number of hours after midnight"""

FMT = '%H:%M:%S'
emails['Time'] = emails['Date'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S').strftime(FMT))
emails['SinceMid'] = emails['Time'].apply(
    lambda x: (datetime.strptime(x, FMT) - datetime.strptime("00:00:00", FMT)).seconds) / 60 / 60

emails.head()

"""Using Wordcloud to see the most used words in the email subjects"""
# Libraries
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# Create a list of words
text = ""
for item in emails["Subject"]:
    if isinstance(item, str):
        text += " " + item
    text.replace("'", "")
    text.replace(",", "")
    text.replace('"', '')

# Create the wordcloud object
wordcloud = WordCloud(width=800, height=800, background_color="white")

# Display the generated image:
wordcloud.generate(text)
plt.figure(figsize=(8, 8))
plt.imshow(wordcloud, interpolation="bilinear")
plt.axis("off")
plt.margins(x=0, y=0)
plt.title("Most Used Subject Words", fontsize=20, ha="center", pad=20)
plt.show()




