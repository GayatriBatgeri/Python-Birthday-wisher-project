import pandas as pd
import datetime
import smtplib
import os
os.chdir(r"C:\Users\Gayatri\PycharmProjects\pythonProject")
# os.mkdir('testing')
#ENTER YOUR CREDENTIAL
GMAIL_ID ='' #enter your mail
#gmail password can be obtained by enabling 2-step varification
GMAIL_PSWD ='' # secure app password
def sendmail(to, sub, msg):
    print(f"Email to {to} sent with subject : {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()


if __name__ == "__main__":
    # sendmail(GMAIL_ID, "subject", "test message")
    # exit()
    df = pd.read_excel("data.xlsx")
    # print(df['Dialogue'])
    today = datetime.datetime.now().strftime("%d-%m")
    # print(type(today))
    yearnow = datetime.datetime.now().strftime("%Y")
    writeInd = []
    for index,item in df.iterrows():
        # print(index, item['Birthday'])
        bday = item['Birthday'].strftime("%d-%m")
        age = item['Birthday'].strftime("%Y")
        # print(bday)
        if today == bday and yearnow not in str(item['Year']):
            sendmail(item['Emailid'], "Happy BirthDay !", f"{item['Dialogue']} happy {int(yearnow) - int(age)}th birthday !!!" )
            writeInd.append(index)
    # print(writeInd)
    if (len(writeInd)!=0):
        for i in writeInd:
            yr = df.loc[i,'Year']
            df.loc[i, 'Year'] = f"{yr} , {yearnow}"
            # print(df.loc[i, 'Year'] )
        # print(df)
        df.to_excel('data.xlsx', index=False)