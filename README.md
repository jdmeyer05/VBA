Excel VBA Code

LOS Exceptions

  - This file was a program I wrote for the Expenditure group for exception analysis. They spent days combing through Excel workbooks manually when I asked them if they wanted me to build something for them. I ended up creating this code. 
  - There is a user form that is also attached to it, but I have not added that here. 
  - The way the code works is the user will select the various features they want to analyze, and the code runs just those portions for them. It also prompts them to input the thresholds they are looking for. 
  - So the program is reasonably flexible, and they are still using it at Sheridan Production currently.
  
Monte Carlo Warhammer 40k
  - Warhammer 40k is a complicated tabletop game. This was just a fun program to create with some valuable visualizations. Basically, it runs a series of randomized tests, collects the data from these tests, which carry through in a sequence, and then outputs a result. I gathered 10000 results and created a histogram. If the results are pretty "Normal," as in normal gaussian distribution, then I calculate a few probabilities to inform the player about the likelihood of the interaction being beneficial or not.
  
Net Volume Analysis Full
  - The main function for me when I was a Revenue Accountant at Sheridan Production was to generate, calculate and analyze the revenue accrual every month. I inherited a process that took several hours and iterations. 
  - I was able to, in the end, get this process down to about ~ the 20s. This includes pulling all the data using SQL code from our database, wrangling the data to make it usable, reconciling the data to make sure it was accurate, organizing the data to be analyzed, performing the analysis, and outputting the results in a way that my manager wanted to see. 
  - Because of this, and some other parts of the process where I did similar optimizations, I could book this accrual in a pinch (i.e., the same day as requested), which came in handy a few times when upper management needed to have a quick understanding of our revenue for the month.
  - Sheridan Production is still using this process currently.
  
PDF Scrape and Character Fix
  - It brings in a messy PDF extraction and cleans it up. This is accomplished by being able to identify several particular issues with the extraction, correcting them using arrays, then outputting the corrected file in a new excel workbook.
