# Lead-Processor
A program I'm actively using and improving as needed to clean and upload sales leads. This program works in tangent with a series of MS Access queries. 

Main functions:
- Run the raw data file through a series of MS Access queries, then download the data
- Perform additional cleaning measures and create validation pivot tables for upload to Salesforce
- Translate these changes to an upload file to an in house database, create validation pivot tables for this file as well
- Separate the files into csvs for upload
- Validate the successful upload and generate a status email on all records
- Archive upload data for trend monitoring
- User friendly GUI with message box wrapped errors for non-programmer users
