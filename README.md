# scraper

By law, pipelines are required to post what is called "Informational Postings" or simply put, data from various times throughout the day.
This data captures the amount of whatever is flowing through the pipeline (Natural Gas Liquids, Crude Oil, Liquefied Natural Gas...etc.) 
The data is entered by "LOC" which is a unique identifier of a certain point along this pipeline.
The quantity we are interested in is called "TSQ" or "Total Scheduled Quantity" which are MMbtu's of whatever is nominated to flow through the pipeline. 

Knowing this information gives us an understanding of the activity on a very in depth level on a certain pipeline, do with the information what you must... 

I have utilized Selenium and BeautifulSoup to open a proxy browser, go to the directed informational posting, simulate mouseclicks as a human would to pull up info.
The scrapers will capture the LOC and the TSQ of that location. 
The specified points are then turned into a DataFrame which creates a burner Excel worksheet.
Information from the burner sheet is appended to a main Excel Database and the burner sheet is then deleted. 

This is to be run daily as the values change daily. Since this is a daily practice, I have added code for counters. In the .txt files, the number present is the row in the excel workbook. 

This project is fully automated. Windows Task Scheduler or Automator (Mac) can be set up to simply execute the script since the counter values are automatically changed to the next row. What is normally a tedious and painstaking process of data collection done once a quarter can be mitigated by running this everyday. 

This project was a part of my data engineering internship.
