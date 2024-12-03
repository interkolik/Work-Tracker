# Work-Tracker
Overview
Work-Tracker is a Python script designed to track work time between the inputs "start" and "stop". After tracking, it asks for a job name and the number of tasks completed, then logs this information.

Features
Work Log Excel File:

Creates a work_log.xlsx file for the current week, dividing the time into 15-minute intervals.

Colors the cells corresponding to the time worked in green (yellow if it overrides anything).

Prints the job name and the number of tasks in the first cell, "done" in the last cell, and "working" in the intermediate cells.

Automatically creates a new sheet each week, pushing older sheets to the right.

Merged Tasks Calendar File:

Creates or merges a merged_tasks.ics file inside the calendar folder to upload into a calendar app.

Merges new events with existing ones, so uploading the .ics file adds all events at once.

Manual Entry Mode:

Includes a manual input mode where start and end times can be entered directly, useful for debugging or manual corrections.

By default, the timezone used is PST.

Installation
Clone this repository to your local machine.

Ensure you have Python installed.

Install the required libraries:

sh
pip install pandas openpyxl pytz
Usage
Run the script:

sh
python work_tracker.py
Follow the prompts to start and stop the timer, and enter the job title and number of tasks completed.

Example
Starting the Timer:

Enter 'start' to begin tracking time and 'stop' to end tracking: start
Started tracking at 2023-12-01 09:00:00
Stopping the Timer:

Enter 'start' to begin tracking time and 'stop' to end tracking: stop
Stopped tracking at 2023-12-01 10:15:00
Enter the job title: Software Development
Enter number of tasks completed: 3
Work time and tasks recorded.
Manual Mode
For manual entry:

Enter 'manual' command:

Enter 'start' to begin tracking time and 'stop' to end tracking: manual
Provide start and end times:

Enter the start time (YYYY-MM-DD HH:MM PST): 2023-12-01 09:00
Enter the end time (YYYY-MM-DD HH:MM PST): 2023-12-01 10:15
Files Created
work_log.xlsx: The weekly work log file.

calendar/merged_tasks.ics: The merged calendar file with tasks.

License
This project is licensed under the MIT License - see the LICENSE file for details.

This code was written with the assistance of Copilot. Feel free to use Copilot to modify or enhance the code as needed.

Contact
If you have any questions or feedback, feel free to reach out!
