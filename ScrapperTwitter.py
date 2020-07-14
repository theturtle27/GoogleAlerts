import GetOldTweets3 as got
import pandas as pd
import openpyxl
from datetime import datetime

keyword = '#istandwithraeesah'
count = 2000

# get date time
now = datetime.now()
date = now.strftime("%d-%m-%y %H%M")

# Creation of query object
tweetCriteria = got.manager.TweetCriteria().setQuerySearch(keyword) \
    .setMaxTweets(count)
# Creation of list that contains all tweets
tweets = got.manager.TweetManager.getTweets(tweetCriteria)
# Creating list of chosen tweet data
text_tweets = [[tweet.date, tweet.text] for tweet in tweets]
print(text_tweets)
# create dataFrame
df = pd.DataFrame(text_tweets, columns=["Time", "Text"])
print(df)
df['Time']=df['Time'].dt.tz_localize(None)
print(df)

# drop duplicates
df.drop_duplicates(subset='Text', inplace=True)
# look for existing file
try:
    wb = openpyxl.load_workbook(keyword + '.xlsx')

    # Inside this context manager, handle everything related to writing new data to the file\
    # without overwriting existing data
    with pd.ExcelWriter(keyword + '.xlsx', engine='openpyxl') as writer:
        # Your loaded workbook is set as the "base of work"
        writer.book = wb

        # Loop through the existing worksheets in the workbook and map each title to\
        # the corresponding worksheet (that is, a dictionary where the keys are the\
        # existing worksheets' names and the values are the actual worksheets)
        writer.sheets = {worksheet.title: worksheet for worksheet in wb.worksheets}

        # Write the new data to the file without overwriting what already exists
        df.to_excel(writer, date, index=False)

        # Save the file
        writer.save()
# write new file
except Exception as e:
    print(e)
    df.to_excel(keyword + '.xlsx',
                sheet_name=date,
                index=False)

# write to csv
df.to_csv(keyword + date + '.csv', index=False)