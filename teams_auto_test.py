import pymsteams

# Create the connectorcard object with the Teams webhook URL
webhook_url = "https://teams.microsoft.com/l/chat/19:67aeb40bb01a4aa0928fc174f0c230b1@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"
teams_message = pymsteams.connectorcard(webhook_url)

# Add text to the message
teams_message.text("此為自動化推播測試")

# Send the message
teams_message.send()