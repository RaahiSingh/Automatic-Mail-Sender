# Automatic Mail Sender

This Python project automates the process of sending emails using Excel roster data.

It reads data from Excel sheets to calculate:

- On Call / OD / Night support days for team members
- Amount payable based on support days
- Name and Signum of each team member

The script automatically generates an HTML email with a summary table and sends it at a scheduled time. Users can also upload Excel files via a simple web interface built with Flask.

## Features

- Reads and processes Excel data
- Generates HTML tables with support summary
- Sends emails automatically using Gmail SMTP
- Scheduler runs daily and sends emails on specific days
- Secure handling of credentials using `.env` file
