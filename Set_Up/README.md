## Municipal Market Data
### Monthly Trade Data Statistics

### Set_Up Folder contains several R scripts to run in succession which extract the data from the downloaded files and into R as data frame tables. Also, the tables are combined into a single monthly data series with raw data that has not been cleaned up. That table is exported locally to a comma delimited (.csv) file are is starting place for the next set of scripts: Clean_and_Enrich



Here is an image of the flow:

Step 1: Download   ===>  Step 2a: Set_up         ==\\
                                                    ===>  Step 3: Clean_and_Enrich
                         Step 2b: Download_More  ==//

 By exporting to .csv files after each step, a user may use a few steps as he or she wants and go a different direction.

*Expected Outputs:*
new_issue_volume.csv
sec_type_volume.csv
_and a third csv to be defined which will have both sets combined_


#### Change History

27 AUG 2016 - K Eiholzer       - Initial R scripts uploaded. Partially _aspirational_, since much of the work is not completed.
