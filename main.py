import pandas as pd
import datetime       # aaj ki date nikalne k lie
import smtplib   
import os

os.chdir(r"C:\Users\subod\OneDrive\Desktop\BirthdayWisher")   # is directory me ana hi pdega coz we are reading data from the xlsx file inside this folder itself.
# os.mkdir("Hola")    # just to check if scheduler is working or not  


GMAIL_ID=""
GMAIL_PASSWORD=""   #  Only work for sending mail — no full access to your Google account

def sendEmail(to , subject , actual_message): 
    print(f"Email sent to {to} with subject: {subject} and the message is : {actual_message}")
    s= smtplib.SMTP("smtp.gmail.com", 587)   # smtp server per entry through port 587 
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PASSWORD)
    s.sendmail(GMAIL_ID, to , f"Subject: {subject} \n\n {actual_message}")    # you must include headers like Subject: manually as part of that string
    s.quit()     




if __name__=="__main__":

    # sendEmail(GMAIL_ID,"HAPPY BIRTHDAY", "Hi Viraat, Happy Bday to you!")      # sirf test krne k lie tha ki mail jaegi ya nhi ye neeche call ho rha h in line 
    # quit()  # just to check that mail is being sent or not

    df=pd.read_excel("data.xlsx")     # excel read and display as df krne k lie bs
    # print(df)   

    today= datetime.datetime.now().strftime("%d-%m")   
    year_now= datetime.datetime.now().strftime("%Y")     # string
    # upar aaj ki date dedega along with exact time like 2025-06-25 15:04:16.312166 and then on that returned value we used strftime to get just the data and month based on the format
    # print(today)                                                                    # type string


    # Time to iterate the dataframe now and get dates in a format %d-%m after fetching today's date and the loading original df from excel

    write_indexes=[]                        # jinko mail ja chuki hai unka index number idhr aa jaega

    for index,series in df.iterrows():                      # it returns a tuple(index,series) which gets unpacked to index and series(variable)
        # print( index , series["Birthdate"])                  # printing the index and the values inside the column Birthdate in each iteration when we write series[column_name]= value  we get one by one and gets printed

        birth_date= series["Birthdate"].strftime("%d-%m")       # ek k bad ek dates aati rhengi idhr variable me and strftime("%d-%m") is a method that belongs to the datetime object got by series["Birthdate"] before using '.' ahead.
        
        #Even though series["Birthdate"] looks like a dictionary access, it returns an object — specifically a datetime object in this case.
        #And because it’s an object, you can call its method with a dot '.'
        # In Python, almost everything is an object .So when you do x[...], x.y, or call a function, you are usually getting an object back — and you can call .method() on it if that object supports it.
        # print(birth_date)  

        if (today==birth_date) and year_now not in str(series["Year"]):         # both are in the same date format  and 2.) checking string in string    
            sendEmail(series["Email"] , "Happy Birthday", series["Dialogue"])
            write_indexes.append(index) 
    print(write_indexes)  
    
    if not write_indexes:
        print("Empty list- No recepient has their birthday today") 



    for i in write_indexes:         # jha jha hmne email bhejdi hai vha ham .loc k through year updated kr rhe hai jisse dobara mail na jaye unnecessarily.
        yr= df.loc[i, "Year"]
        print(yr) 
        df.loc[i,"Year"]= str(yr) + "," + str(year_now)   # df is mutable and the value can be changed using loc in place (value can be access and changed using .loc() )

        # print(df.loc[i,"Year"])
    # print(df)

    df.to_excel("data.xlsx",index=False)      