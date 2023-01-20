# ExcelDataFrame
 pandas.DataFrame subclass links with Excel

# Usage
 Almost all is same as pandas.DataFrame but push() to send data to Excel and pull() to receive data from Excel.
 
 ```py
 df = DataFrame(data, columns=columns)
 df.push() # -> open Excel and show data.
 df.pull() # -> pull data from Sheet1 of an activated Excel.
 
 # push() returns data then:
 df = df.apply(lambda x: 2*x).push()
 ```
