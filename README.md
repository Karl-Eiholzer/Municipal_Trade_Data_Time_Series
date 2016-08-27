## Municipal Market Data
### Monthly Trade Data Statistics

### Contains R Scripts needed to download and perform initial structuring and cleaning of MS Excel spreadsheets published by Municipal Securities Rulemaking Board. This unique set of spreadsheets contain transaction volume published continuously since 2000.

This Repository is divided into sub folders:
 - Documentation - reference material to understand the data
 - Download - R Script designed to move the full set of spreadsheets from the MSRB's server to your local drive. These local files are the starting place for the next scripts: Set_up.
 - Set_Up - Several R scripts to run in succession which extract the data from the downloaded files and into R as data frame tables. Also, the tables are combined into a single monthly data series with raw data that has not been cleaned up. That table is exported locally to a comma delimited (.csv) file are is starting place for the next set of scripts: Clean_and_Enrich
 - Download_More -  Additional sources of data which are specific to the municipal space are downloaded and merged into a monthly series. This table is exported locally to a comma delimited (.csv) file and is starting place for the next set of scripts: Clean_and_Enrich
 - Clean_and_Enrich - Starts with the comma delimited files created by the Set-Up and Download_More scripts by merging them into a single monthly time series. The data then has minor cleaning and two comma delimited files are created: one with clean, actual data and a second with the same data transformed into normalized ranges (to support exploratory analysis).

Here is an image of the flow:

Step 1: Download   ===>  Step 2a: Set_up         ==\\
                                                    ===>  Step 3: Clean_and_Enrich
                         Step 2b: Download_More  ==//

 By exporting to .csv files after each step, a user may use a few steps as he or she wants and go a different direction.



#### Change History

27 AUG 2016 - K Eiholzer       - Described plan for repository. Partially _aspirational_, since much of the work is not completed.
