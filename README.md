# KQL_OutsideWorkingHours
Powershell script to generate KQL (Keyword Query Language). Used in Purview eDiscovery Review Sets, this enables Reviewers and Investigators to target content outside specified working hours.

Execute by navigating to the .ps1 and running the script with .\ (fullstop backslash).
KQL will be generated 1 month at a time - enter the required timeframe in MM-YYYY format. 
Output will be stored as .txt file in same location as the .ps1, ready to copy into Purview Review Set under 'KQL' option.
Example -
<img width="565" height="49" alt="image" src="https://github.com/user-attachments/assets/35d6ba90-db64-4989-b3ac-c125060734de" />

For background see - https://purviewpower.wordpress.com/2025/12/19/outside-working-hours-kql-filter-for-ediscovery-review-sets/
