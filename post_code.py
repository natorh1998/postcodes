'''
This assumes that you have basic knowledge of Python 3 but limited knowledge of particular packages and concepts as I don't know how advanced you are currently. The first section includes a regex for parsing
post codes (in case they're not in the nice valid format) and the second applies this to searching for the relevant region.
'''

#Import the relevant libraries using the standard notation
import numpy as np
import re
import pandas as pd

#Compile the regex to be used in post code parsing.
'''
Each section does the following:
[A-Z]{1,2} Specifies the uppercase alphabet for a minimum of one and maximum of two characters
\d[A-Z\d]? specifies any digit, then the uppercase alphabet and any digit. Including it in [] means that it is searched as a group (ie. any one of those). The ? operator is a lazy operator that specifies once or none.
Then the space and ? specifies that there may be a space once or none times.
Finally, the last portion specifies a digit then the uppercase alphabet two times.

This ensures that when you have a messy post codes variable, or post codes included at any point in a string, it will pick up any pattern of post codes.
For example,
NG71LR
NG7 1LR
NG171LR
NG17 1LR

can all be read into valid postcodes.
'''
p = re.compile('[A-Z]{1,2} ?\d[A-Z\d]? ?\d[A-Z]{2}')

#Set the locations of the post codes file and the data file.
post_loc = 'postcodes.csv' #this is in my C: drive and hence the same location as the anaconda root directory, meaning it doesn't need to be specified in full
data_loc = r'C:\Users\norme-herbert\OneDrive - Department for Education\Downloads\PAWR NH Employees and Home Postcodes 2020-10-08 08_59 BST.xlsx' #the r operator in front of the string tells python to automatically escape special characters, in this case the \ character.


post_codes = pd.read_csv(post_loc) #import the post codes file using pandas
post_codes['town'] = post_codes['town'].fillna(post_codes['region']) #fill missing values in the town column with values from the region column

df = pd.read_excel(data_loc, skiprows=list(range(4))) #read in the data file. As the file column headers don't start until row 5, we use skiprows to set that. Skiprows takes a list-like variable so we create a range of 0-3 then convert it to a list
#df['Primary Home Address - Postal Code'] = df['Primary Home Address - Postal Code'].str.replace(" ", "")#Remove whitespace from postcodes. We could also do this in regex, but given the combinations it's easier to do it here

#Create a new column in the dataframe called Town. The apply method allows us to apply a function to every cell in a column (or full dataframe). This uses a lambda method to reduce the length of code we write.
#This takes the post code, parses it to ensure it is correct, splits on the space character and selects the first portion (eg. NG7)
#It then filters the postcodes database to find the postcode which matches the parsing, then selects the first row in the 'town' column
#if the parser finds no valid postcode, the function returns null
df['Town'] = df['Primary Home Address - Postal Code'].astype(str).apply(lambda x: "Null" if (p.search(x) == None) else (post_codes[post_codes['postcode'] == p.search(x).group().split(" ")[0]].iloc[0]['town'] if (len(post_codes[post_codes['postcode'] == p.search(x).group().split(" ")[0]]) > 0) else "Invalid"))

#As before, but this time taking the first row of the region column as opposed to the town column
df['Region'] = df['Primary Home Address - Postal Code'].astype(str).apply(lambda x: "Null" if (p.search(x) == None) else (post_codes[post_codes['postcode'] == p.search(x).group().split(" ")[0]].iloc[0]['region'] if (len(post_codes[post_codes['postcode'] == p.search(x).group().split(" ")[0]]) > 0) else "Invalid"))

#Again, but wider regions
df['UK Region'] = df['Primary Home Address - Postal Code'].astype(str).apply(lambda x: "Null" if (p.search(x) == None) else (post_codes[post_codes['postcode'] == p.search(x).group().split(" ")[0]].iloc[0]['uk_region'] if (len(post_codes[post_codes['postcode'] == p.search(x).group().split(" ")[0]]) > 0) else "Invalid"))

#export to excel
df.to_excel('region_output.xlsx', index=False)