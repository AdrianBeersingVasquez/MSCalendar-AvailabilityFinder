# MSCalendar-AvailabilityFinder

MSCalendar-AvailabilityFinder is a Python script that connects to the Microsoft Graph API to read shared or personal calendars, in order to identify time slots over the ensuing week when there are no scheduled events. This tool is ideal for users who need to quickly find free time slots in a busy schedule. The script returns a text output that can be copied and pasted into a draft email, to assist in scheduling meetings.

## Features

- **Authentication with Microsoft Graph API**: Securely access your calendar data using OAuth2.
- **Time Zone Conversion**: Convert event times to a specified time zone for easier scheduling.
- **Event Filtering**: Automatically filter out events that fall outside of sender's and recipient's working hours.
- **Weekend Filtering**: Remove weekend slots for both sender and receiver.
- **Fancy Printing**: Print the schedule in a readable format.

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/MSCalendar-AvailabilityFinder.git
    cd MSCalendar-AvailabilityFinder
    ```

2. Install the required dependencies:
    ```sh
    pip install O365 pandas
    ```

## Usage

1. **Setup Authentication Credentials**:
    Replace the `CLIENT_ID` and `SECRET_ID` in the script with your Microsoft Graph API credentials.
    ```python
    CLIENT_ID = 'your-client-id'
    SECRET_ID = 'your-secret-id'
    credentials = (CLIENT_ID, SECRET_ID)
    ```

2. **Define Time Range**:
    Specify the start and end dates for fetching events:
    ```python
    start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=timezone.utc) + timedelta(days=1)
    end_date = start_date + timedelta(days=5)
    ```

3. **Enter Email Address of the Calendar of Interest**:
    Replace `'email@example.com'` with the email address of the person whose calendar you are scheduling for:
    ```python
    schedule = account.schedule(resource='email@example.com')
    ```

4. **Display the Schedule**:
    Print the formatted schedule with available time slots:
    ```python
    fancy_print(availability_df)
    ```

## Notes

- Ensure to modify the credentials and time ranges as per your requirements.
- This script is currently configured to convert to the the **America/New_York** time zone, but it can be easily adapted to other time zones by changing **client_timezone**.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.
