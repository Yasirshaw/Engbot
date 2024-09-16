import pandas as pd
import telebot
import schedule
import time
import logging
import os
from datetime import datetime
import pytz

# Initialize logging
logging.basicConfig(level=logging.INFO)

# Initialize the Telegram bot with your token
API_TOKEN = '7364384207:AAGFQEJ6-SdGoZSuSc3WATxwZjDnBlgFHj8'
bot = telebot.TeleBot(API_TOKEN)

# Use your actual channel ID
channel_id = '-1001798059502'

# Define the timezone for Saudi Arabia
saudi_tz = pytz.timezone('Asia/Riyadh')

# File to track the last sent word index
tracking_file = 'data/last_word_index.txt'

# Function to load the words and definitions from the Excel sheet
def load_words_from_excel():
    try:
        df = pd.read_excel('data/words.xlsx')  # Ensure your Excel file is in the right location
        return df
    except FileNotFoundError:
        logging.error("The Excel file was not found.")
        return pd.DataFrame()
    except Exception as e:
        logging.error(f"An error occurred while loading the Excel file: {e}")
        return pd.DataFrame()

# Function to get the next word index
def get_next_word_index():
    if not os.path.exists(tracking_file):
        return 0  # Start from the first word if the file does not exist

    with open(tracking_file, 'r') as file:
        try:
            index = int(file.read().strip())
        except ValueError:
            index = 0  # Default to 0 if there's an issue with reading the file

    return index

# Function to update the last sent word index
def update_last_word_index(index):
    with open(tracking_file, 'w') as file:
        file.write(str(index))

# Function to send the word of the day from the Excel sheet
def send_word_of_the_day():
    words_df = load_words_from_excel()

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

# Schedule the bot to send the word of the day at 4 PM Saudi Arabia time
schedule.every().day.at("09:00").do(send_word_of_the_day)

# Polling and scheduling loop
while True:
    # Check for scheduled tasks (like sending at 4 PM Saudi time)
    schedule.run_pending()

    # Polling for manual command triggers
    bot.polling(none_stop=True)

    # Prevent CPU overload
    time.sleep(60)  # Check every minute for scheduled tasks
