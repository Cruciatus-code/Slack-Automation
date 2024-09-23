__author__ = "Lavisha"
__version__ = "2.0"
__description__ = ("A Python program to extract case information from a specifc Slack channel")


import os
import pandas as pd
import pytz
from slack_sdk import WebClient #type:ignore
from slack_sdk.errors import SlackApiError  #type:ignore
from datetime import datetime , timezone

# Set your Slack API token
SLACK_TOKEN = "your_Slack_token_here"
client = WebClient(token=SLACK_TOKEN)

# Function to retrieve messages from a channel with pagination support
def fetch_channel_messages(channel_id, limit=1000):
    messages = []
    cursor = None

    try:
        while True:
            result = client.conversations_history(channel=channel_id, limit=limit, cursor=cursor)
            messages.extend(result['messages'])

            # Get the next cursor, if there is more data
            cursor = result.get('response_metadata', {}).get('next_cursor')
            if not cursor:
                break

    except SlackApiError as e:
        print(f"Error fetching messages: {e}")
    
    return messages



# Convert timestamp to a specific timezone and format (DD-MM-YYYY HH:MM:SS)
def format_timestamp(ts, timezone_str='UTC'):
    utc_time = datetime.fromtimestamp(float(ts), tz=timezone.utc)
    local_time = utc_time.astimezone(pytz.timezone(timezone_str))
    return local_time.strftime("%d-%m-%Y %H:%M:%S")

# Convert DD-MM-YYYY HH:MM:SS to Unix timestamp (for start and end date)
def convert_to_timestamp(date_str):
    try:
        dt_object = datetime.strptime(date_str, "%d-%m-%Y %H:%M:%S")
        return dt_object.replace(tzinfo=timezone.utc).timestamp()
    except ValueError:
        print(f"Invalid date format: {date_str}")
        return None

# Function to check for case-related messages
def find_case_posts (messages, start_ts, end_ts):
    case_posts = []
    for message in messages:
        text = message.get('text', '')
        message_ts = float(message.get('ts'))

        # Check if message timestamp is within the range
        if start_ts <= message_ts <= end_ts:
            case_number = find_case_or_ctask_number(text, 'CS')
            ctask_number = find_case_or_ctask_number(text , 'CTASK')
            

            if case_number or ctask_number:
                case_description = find_case_description(text)
                post_date = format_timestamp(message.get('ts'), 'Asia/Kolkata')  # Date in DD-MM-YYYY (IST) format
                case_posts.append({
                    'Case Number': case_number,
                    'Ctask Number': ctask_number,
                    'Case Description': case_description,
                    'Post Date (IST) ': post_date
                })
    return case_posts

# Find case and task numbers using a reusable pattern function
def find_case_or_ctask_number(text, prefix):
    for word in text.split():
        parts = word.split('/')
        for part in parts:
            if part.startswith(prefix):
                return part
    return None


# Function to extract case description
def find_case_description(text):
    return text  

# Function to save the case posts to an Excel file
def save_to_excel(case_posts, filename="Extract_case_posts.xlsx"):
    df = pd.DataFrame(case_posts)
    df.to_excel(filename, index=False)
    print(f"\nData saved to {filename}")

    
# Main function to get case posts from a Slack channel within a specified date range
def get_case_posts(channel_id):
    # Ask user for start and end dates
    start_date = input("Enter the start date and time (DD-MM-YYYY HH:MM:SS): ")
    end_date = input("Enter the end date and time (DD-MM-YYYY HH:MM:SS): ")
    
    # Convert the start and end dates to Unix timestamps
    start_ts = convert_to_timestamp(start_date)
    end_ts =  convert_to_timestamp(end_date)

    if start_ts is None or end_ts is None:
        print("Invalid dates provided. Please try again.")
        return
    
    # Fetch all messages
    messages = fetch_channel_messages(channel_id)
    
    # Filter and find case-related posts within the date range
    case_posts = find_case_posts(messages, start_ts, end_ts)
    number_of_case_posts = len(case_posts)
    
    if number_of_case_posts > 0:
        print(f"\nNumber of case/CTask posts found: {number_of_case_posts}")
        
        # Save the case posts to an Excel file
        save_to_excel(case_posts)
    else:
        print("\nNo case/CTask posts found within the specified date range.")

# Example usage
channel_id = "your_channel_id_here"  # Replace with your channel ID
get_case_posts(channel_id)

