# Zoom Meeting Scheduler

A Python application that automatically schedules Zoom meetings from Excel data with GUI file selection. The script reads meeting information from an Excel file and creates scheduled Zoom meetings via the Zoom API, with built-in duplicate detection and multi-language support.

## Features

- **Excel Integration**: Read meeting data directly from Excel files (.xlsx) with structured format
- **GUI File Selection**: User-friendly file picker using Tkinter
- **Duplicate Detection**: Automatically checks for existing meetings to avoid duplicates
- **Multi-day Scheduling**: Schedule meetings for up to two different days of the week per row
- **Flexible Scheduling**: Supports various day combinations (Monday/Wednesday, Tuesday/Thursday, etc.)
- **Timezone Support**: Handles GMT and European time zones with automatic conversion
- **Unicode Support**: Full support for special characters and emojis in meeting content
- **Detailed Logging**: Generates formatted meeting information with join links and instructions
- **Smart Date Calculation**: Automatically calculates next occurrence of specified weekdays
- **Batch Processing**: Process multiple meeting entries from a single Excel file

## Prerequisites

### Python Dependencies
```bash
pip install requests openpyxl
```

### Required Libraries
- `requests` - For Zoom API communication
- `openpyxl` - For reading Excel files
- `tkinter` - For GUI file selection (usually included with Python)
- `json` - For API data handling (built-in)
- `datetime` - For date/time operations (built-in)
- `io` - For UTF-8 file handling (built-in)

### Zoom API Credentials
You'll need to create a Zoom App and obtain:
- Account ID
- Client ID
- Client Secret

## Setup

### 1. Zoom API Setup
1. Go to [Zoom Marketplace](https://marketplace.zoom.us/)
2. Sign in and click "Develop" → "Build App"
3. Choose "Server-to-Server OAuth" app type
4. Fill in the required information
5. Get your Account ID, Client ID, and Client Secret
6. Add the following scopes:
   - `meeting:write:admin`
   - `meeting:read:admin`
   - `user:read:admin`

### 2. Script Configuration
Edit the script and add your Zoom API credentials:
```python
ACCOUNT_ID = 'your_account_id_here'
CLIENT_ID = 'your_client_id_here'
CLIENT_SECRET = 'your_client_secret_here'
```

### 3. Excel File Format
Your Excel file should have the following structure with headers in row 1:

| Column A | Column B | Column C | Column D | Column E | Column F | Column G |
|----------|----------|----------|----------|----------|----------|----------|
| **Name** | **Email** | **Subject** | **Day 1** | **Day 2** | **Hour** | **Minute** |
| John Deb | abc@gmail.com | Dev team meeting | Monday | Wednesday | 19 | 30 |
| Arthor See | abd@gmail.com | Compensation | Monday | Wednesday | 21 | 0 |
| Carlos Jr | kdb@gmail.com | Interview | Monday | Wednesday | 20 | 0 |
| Marcos Ty | mcr@gmail.com | IDK | Monday | Wednesday | 20 | 30 |

**Column Details:**
- **Column A (Name)**: Host name (informational only)
- **Column B (Email)**: Zoom account email for meeting host
- **Column C (Subject)**: Meeting topic/title
- **Column D (Day 1)**: First meeting day (Monday, Tuesday, Wednesday, etc.)
- **Column E (Day 2)**: Second meeting day (can be empty if only one day needed)
- **Column F (Hour)**: Meeting hour in 24-hour format (0-23)
- **Column G (Minute)**: Meeting minute (0-59)

**Important Notes:**
- The script starts reading from row 2 (row 1 should contain headers)
- Email addresses must be valid Zoom account emails
- Day names must be spelled correctly (Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday)
- Time is processed in 24-hour format
- If only one meeting day is needed, leave Day 2 column empty

## Usage

### Running the Script
1. Run the Python script:
   ```bash
   python zoom_scheduler.py
   ```

2. A file dialog will appear - select your Excel file

3. The script will:
   - Read the Excel data
   - Check for existing meetings
   - Schedule new meetings as needed
   - Generate a `meetings.txt` file with meeting details

### Output
The script generates a `meetings.txt` file containing:
- Meeting join URLs
- Meeting IDs
- Host email addresses
- Formatted meeting times (GMT and European)
- Instructions in French for participants

### Example Use Cases

**Language Learning Classes:**
- Schedule recurring language classes for different instructors
- Multiple time slots per week (e.g., Monday/Wednesday, Tuesday/Thursday)
- Different subjects (MW-ABG, MW-IG, TT-BBG representing different course types)

## How It Works

### 1. File Processing
- Uses `openpyxl` to read Excel files
- Processes each row starting from row 2
- Extracts email, topic, days, and time information

### 2. Date Calculation
- Calculates the next occurrence of specified weekdays
- Handles cases where the meeting day is today but time hasn't passed
- Supports scheduling for multiple days per week

### 3. Zoom API Integration
- Authenticates using Server-to-Server OAuth
- Checks for existing meetings to prevent duplicates
- Creates scheduled meetings with specified parameters

### 4. Meeting Scheduling
- Creates meetings with 60-minute default duration
- Sets timezone to UTC
- Generates detailed meeting information

### 5. Output Generation
- Creates formatted meeting announcements
- Includes multilingual instructions
- Provides all necessary joining information

## File Structure
```
zoom-scheduler/
├── zoom_scheduler.py      # Main script
├── meetings.txt           # Generated output (after running)
├── requirements.txt       # Python dependencies
└── README.md             # This file
```

## Configuration Options

### Modifying Meeting Duration
Change the duration in the `process_excel_data` function:
```python
schedule_meeting(email, topic, meeting_datetime, 90, log_file)  # 90 minutes
```

### Changing Timezone
Modify the timezone in the `schedule_meeting` function:
```python
data = {
    'topic': topic,
    'type': 2,
    'start_time': start_time_str,
    'duration': duration,
    'timezone': 'Europe/Paris'  # Change timezone here
}
```

### Customizing Output Format
Edit the text formatting in the `schedule_meeting` function to customize the meeting announcement format.

## Error Handling

The script includes error handling for:
- **Authentication failures**: Invalid API credentials
- **File access errors**: Missing or corrupted Excel files
- **API rate limits**: Zoom API request limitations
- **Unicode encoding**: Special characters in meeting content
- **Duplicate meetings**: Existing meetings with same topic and time

## Troubleshooting

### Common Issues

**UnicodeEncodeError**
- Ensure the script uses UTF-8 encoding for file operations
- Check that special characters are properly handled

**Authentication Error**
- Verify your Zoom API credentials
- Ensure your Zoom app has the required scopes
- Check that your app is activated

**Excel Reading Error**
- Verify the Excel file format matches the expected structure (7 columns: Coach, Email, Subject, Day 1, Day 2, Hour, Minute)
- Ensure the file is not corrupted or password-protected
- Check that the file has data starting from row 2 with headers in row 1
- Verify day names are spelled correctly (Monday, Tuesday, etc.)
- Ensure hour values are in 24-hour format (0-23) and minute values are 0-59

**Meeting Creation Failure**
- Verify the email addresses are valid Zoom users
- Check that meeting times are in the future
- Ensure the host has permission to schedule meetings

## Security Notes

- Keep your API credentials secure and never commit them to version control
- Consider using environment variables for sensitive information
- Regularly rotate your API credentials
- Limit API app permissions to only what's necessary

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source. Please check the license file for details.

## Support

For issues related to:
- **Zoom API**: Check [Zoom API Documentation](https://developers.zoom.us/docs/api/)
- **Python dependencies**: Refer to individual package documentation
- **Script functionality**: Open an issue in this repository

## Version History

- **v1.0**: Initial release with basic scheduling functionality
- **v1.1**: Added duplicate detection and improved error handling
- **v1.2**: Enhanced Unicode support and GUI file selection
- **v1.3**: Improved date calculation and timezone handling

---

**Note**: This script is designed for educational and productivity purposes. Ensure compliance with your organization's policies and Zoom's terms of service when using automated meeting scheduling.
