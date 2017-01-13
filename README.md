This solution scripts objects from SQL Server into text files and commits it into GitHub repository.
The idea is to have some level of change tracking without using complicated and expensive tools for source control.
Objects to script:
- on SQL Server level
  * SQL Jobs
  * SQL Logins
  * configuration parameters
- on database level
  * all objects: tables, views, sprocs, functions...
  * users and permissions
