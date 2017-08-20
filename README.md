# OrganizeReports
Python script used to organize 5000+ safety reports for a construction company I worked for. No reports included only the code.

Begins with a large list of randomly ordered safety reports, then creates a system of folders organized by subcompany and projects within that subcompany.

Then uses openpyxl to create an excel sheet for each project that contains a list of all safety reports. Hopefully will be updated with a bot to download detail reports from the web.

Project finished 8/16. downloadReports.py is a selenium bot that traverses the safety website for detail reports (weekly updates for each job) for each construction job, then downloads them and puts them in the appropriate file. Very very slow due to the nature of the website, but it gets the job done. Overrall very unoptimized.
