# Automated Maintenance Reminder
## Overview
This program came to be from the need for a reporter and a logger to govern hardware fixtures maintenance. The program scans through an excel spreadsheet, logging each unique 
serial number corresponding to a fixture and its latest maintenance date. The program then emails that report on a scheduled job to a few specific staff members to address
the maintenance of the fixtures. Information provided in the email for each fixture include:
1. Part number
1. Serial number
1. Date of last maintenance
1. Time elapsed since last maintenance
1. Next maintenance and time left

The fixtures are to be inspected on a 16-week (4 months) cycle, for a total of 36 fixtures.

## Technicalities
The program runs from the main class Program.cs which calls the helper classes FixturesChecker.cs and Emailer.cs. FixturesChecker does the main work of scanning through the
spreadsheet and recording all maintenance dates for a specific fixture. Those maintenance dates are then filtered for the latest date which is then set as the official last
maintenance date. The fixtures are stored in a Hashset to ensure no duplicates are present. Once a maintenance date is identified, several checks are performed to deduce how close
the fixture is to its maintenance due date. 

There are three levels of urgency. Green indicates next maintenance over 30 days away, yellow is between 30 days and 14 days while red is anything less than 14 days away. These
three levels subsequently dictate the frequency of the emails. The scheduled job runs daily at 9AM. However, emails are only sent out depending on the state of urgency flagged
by the program. If a green flag is indicated, an email is sent out once a week on Monday. If a yellow flag is indicated, an email is sent out twice a week. For a red flag,
emails are sent out daily at 9AM until the specific fixtures are addressed and updated in the spreadsheet.

The emailer uses an SMTP relay for sending out the email internally.
