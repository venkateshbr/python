
# coding: utf-8

# In[21]:


import urllib
import urllib.request


# In[22]:


urlofFilename = "https://www.nseindia.com/content/historical/EQUITIES/2018/JUN/cm05JUN2018bhav.csv.zip"


# In[23]:


urlofFilename


# In[24]:


localzipfilePath = "C:\\Dev\\LearningPython\\cm05JUN2018bhav.csv.zip"


# In[25]:


hdr = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.62 Safari/537.36'
       }


# In[26]:


hdr


# In[27]:


webreqest = urllib.request.Request(urlofFilename, headers=hdr)


# In[29]:


try:
    page = urllib.request.urlopen(webreqest)
    content = page.read()
    output = open(localzipfilePath, "wb")
    output.write(bytearray(content))
    output.close()
except(urllib.request.HTTPError) as e:
    print(e.fp.read())
    print("Looks like the file download did not happen for url = ", urlofFilename)


# In[30]:


import zipfile
import os


# In[34]:


localExtractPath = "C:\\Dev\\LearningPython"


# In[44]:


if os.path.exists(localzipfilePath):
    print('Cool!' + localzipfilePath + " exists .. extracting")
    listofFiles = []
    fh = open(localzipfilePath, 'rb')
    zipfileHandler = zipfile.ZipFile(fh)
    for filename in zipfileHandler.namelist():
        zipfileHandler.extract(filename, localExtractPath)
        listofFiles.append(localExtractPath+"\\"+filename)
        print("Extracted " + filename +
              " from the zip file to " + localExtractPath+filename)
    print("In total, we extracted ", str(len(listofFiles)), " files")
    fh.close()


# In[45]:


import csv


# In[49]:


onefileName = listofFiles[0]


# In[72]:


lineNum = 0


# In[73]:


listofLists = []


# In[74]:


with open(onefileName, 'r') as csvfile:
    lineReader = csv.reader(csvfile, delimiter=",", quotechar="\"")
    for row in lineReader:
        lineNum = lineNum + 1
        if (lineNum == 1):
            print("Skipping the Header Row")
            continue
        symbol = row[0]
        close = row[5]
        prevClose = row[7]
        tradeQty = row[9]
        pctChange = float(close)/float(prevClose) - 1
        oneResultRow = [symbol, pctChange, float(tradeQty)]
        listofLists.append(oneResultRow)
        print(symbol + " " + "{:,.1f}".format(float(tradeQty) /
                                              1e6) + "M INR", "{:,.1f}".format(pctChange*100)+"%")
    print("Done iterating over the file contents. the file is closed now")
    print("we have stock info for " + str(len(listofLists)) + " stocks")


# In[75]:


listoflistsSortedbyQty = sorted(listofLists, key=lambda x: x[2], reverse=True)


# In[76]:


listoflistsSortedbyQty = sorted(listofLists, key=lambda x: x[1], reverse=True)


# In[77]:


# In[78]:


listoflistsSortedbyQty = sorted(listofLists, key=lambda x: x[2], reverse=True)


# In[79]:


# In[82]:


import xlsxwriter


# In[83]:


excelFilename = "C:\\Dev\\LearningPython\\cm05JUN2018bhav.xlsx"


# In[84]:


workbook = xlsxwriter.Workbook(excelFilename)


# In[85]:


worksheet = workbook.add_worksheet("Summary")


# In[86]:


worksheet.write_row("A1", ["Top traded Stocks"])


# In[87]:


worksheet.write_row("A2", ["Stock", "% Change", "Value Traded (INR)"])


# In[88]:


for rowNum in range(5):
    oneRowtoWrite = listoflistsSortedbyQty[rowNum]
    worksheet.write_row("A" + str(rowNum + 3), oneRowtoWrite)
workbook.close()

print("The Summary of top 5 stocks by volume in Excel " +
      excelFilename + " is generated")
