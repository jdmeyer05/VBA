# VBA-Code
Excel VBA Code

LOS Exceptions
  - This file was a program I wrote for the Expenditure group for exception analysis. They were spending days combing through Excel workbooks manually when I asked them if they wanted me to build something for them. I ended up building this code. There is a userform that is also attached to it but I have not added that here. The way this works is the user will select the various features they want to analyze and the code runs just those portions for them. It also prompts them to input w/e thresholds they are looking for. So the program is fairly flexible and they are still using it at Sheridan Production to this day.

Monte Carlo Warhammer 40k
  - Warhammer 40k is a complicated table top game. This was just a fun program to create with some useful vizualizations. Basically it runs a series of randomized tests, collects the data from these tests which carry through in a sequence and then outputs a result. I gather 10000 results and create a historgram. If the results are fairly "Normal" as in normal gaussian distribution then I calculate a few probabilites to inform the player about the likelihood of the interaction being beneficial or not.

Net Volume Analysis Full
  - A main function for me when I was a Revenue Accountant at Sheridan Production was to generate, calculate and analyze the revenue accrual every month. I adopted a process that took several hours and interations. I was able to, in the end, get this process down to about ~20s. This includes pulling all the data using SQL code from our database, wrangling the data to make it usable, reconciliting the data to make sure it was accurate, organizing the data so be analyzed, perform the analysis and output the results in a way that my manager wanted to see. Because of this, and some other parts of the process where I did similar optimizations, I could book this accrual in a pinch (i.e. same day as requested) which came in handy a few times when upper management needed to have a quick understanding of our revenue for the month.

PDF Scrape and Character Fix
  - Brings in a messy PDF extraction and cleans it up. This is accomplished by being able to identify several particular issues with the extraction, correcting them using arrays, then ouputting the corrected file in a new excel workbook
