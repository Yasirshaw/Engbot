import pandas as pd
import telebot
from apscheduler.schedulers.background import BackgroundScheduler
import logging
import os
from datetime import datetime
import pytz

# Initialize logging
logging.basicConfig(level=logging.WARNING)  # Reduce logging level

# Initialize the Telegram bot with your token
API_TOKEN = '7364384207:AAGFQEJ6-SdGoZSuSc3WATxwZjDnBlgFHj8'  # Replace with your actual bot token
bot = telebot.TeleBot(API_TOKEN)

# Use your actual channel ID
channel_id = -1001798059502  # Replace with your actual channel ID

# Define the timezone for Saudi Arabia
saudi_tz = pytz.timezone('Asia/Riyadh')

# Load the words DataFrame once at startup
words_df = pd.read_excel('data/words.xlsx')  # Ensure your Excel file is in the right location
current_index = 0

# Function to get the next word index from file
def get_next_word_index():
    global current_index
    try:
        with open('data/last_word_index.txt', 'r') as file:
            current_index = int(file.read().strip())
    except (FileNotFoundError, ValueError):
        current_index = 0
    return current_index

# Function to update the last sent word index in file
def update_last_word_index(index):
    global current_index
    current_index = index
    with open('data/last_word_index.txt', 'w') as file:
        file.write(str(index))

# Function to send the word of the day from the Excel sheet
def send_word_of_the_day():
    global current_index
    if words_df.empty:
        logging.warning("The words DataFrame is empty. No message sent.")
        return

    index = get_next_word_index()
    if index >= len(words_df):
        index = 0  # Reset to the first word if index is out of bounds

    # Get the word at the current index
    word = words_df.iloc[index]['Word']
    arabic_def = words_df.iloc[index]['Arabic Definition']
    example_sentence = words_df.iloc[index]['Example Sentence']

    # Update the index for the next day
    update_last_word_index(index + 1)

    # Craft the message with improved spacing and Markdown formatting
    message = (
        f"🎓 *Word of the Day* | *كلمة اليوم*\n\n"
        f"**{word}** - **{arabic_def}**\n\n"
        f"📖 *Example | مثال*: \n\n"
        f"{example_sentence}\n\n"
        f"💡 *Try using the word in a sentence today!*\n"
        f"💡 *حاول استخدام الكلمة في جملة اليوم!* 😊"
    )

    # Send the message to your private channel
    try:
        bot.send_message(chat_id=channel_id, text=message, parse_mode='Markdown')
        logging.info("Word of the Day message sent successfully to the channel.")
    except Exception as e:
        logging.error(f"An error occurred while sending the message: {e}")

# Function to manually trigger the word of the day via command
@bot.message_handler(commands=['send_word'])
def manual_word_of_the_day(message):
    send_word_of_the_day()

# Initialize the scheduler
scheduler = BackgroundScheduler(timezone=saudi_tz)
scheduler.add_job(send_word_of_the_day, 'cron', hour=11, minute=25)
scheduler.start()

# Start polling for bot commands
bot.polling(none_stop=True)