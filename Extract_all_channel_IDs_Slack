from slack_sdk import WebClient #type:ignore
from slack_sdk.errors import SlackApiError #type:ignore


# Your Slack Bot Token
SLACK_TOKEN = "your_slack_token_here"

client = WebClient(token=SLACK_TOKEN)

try:
    # Call the conversations_list method to retrieve all channels
    response = client.conversations_list()

    channels = response["channels"]
    for channel in channels:
        print(f"Channel Name: {channel['name']}, Channel ID: {channel['id']}")

except SlackApiError as e:
    print(f"Error fetching channels: {e}")
