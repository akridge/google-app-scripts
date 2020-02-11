# Master CALENDAR SCRIPT

***

## Table of Contents
1. **[Description](#description)**
2. **[Contact](#contact-developer)**
3. **[Setup](#setup)**

# Description

This script is designed to update a master calendar with individual calendar events from multi-user calendars. 

Events are only copied when they contain one of the following tags:
*  Leave (i.e. Annaul Leave, Sick Leave, Family Leave)
*  Sick
*  Vacation
*  AWS
*  Flex
*  Telework
*  Offsite(Off-site)

The above tags are grouped as follows:
*  Leave: Leave, Sick, Vacation
*  AWS: AWS, Flex
*  Telework: Telework
*  Off-site; Off-site, Offsite

***

# Setup

1. Define the master calendar Id in the "MASTER_CALENDAR_SCRIPT.gs" file to populate the master calendar
2. Define the email_list google folder id below 
3. Populate export.csv.xlsx google sheet with emails
4. Setup time-based triggers:
*  a. Calendars must be shared with the user who is running the script in order to capture the events
*  b. setup triggers to run script at certain times. 1am and 10am

