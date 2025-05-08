# GoogleChats
Python3. Intakes a directory where Google Chats are stored from a search warrant return from Google. Typically (after extracting the .zips), this directory will be akin to '...GoogleChat.Messages_001.001/Google Chat/Groups/". Will walk all sub-folders of the directory and create a single excel timeline for all messages with chatID, datetime UTC, sender email address, body, attachment (hyperlinked), and IP address.

# Requirements
- pip install openpyxl

# Usage
`googleChats.py <parentDirectory>`
 
 googleChats.xlsx will be created in the <parentDirectory> passed.
