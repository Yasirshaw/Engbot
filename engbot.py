import pandas as pd
import telebot
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.executors.pool import ThreadPoolExecutor
import logging
import os
from datetime import datetime, timedelta
import pytz

# Initialize logging
logging.basicConfig(level=logging.INFO)

# Initialize the Telegram bot with your token
API_TOKEN = '7364384207:AAGFQEJ6-SdGoZSuSc3WATxwZjDnBlgFHj8'  # Replace with your actual bot token
bot = telebot.TeleBot(API_TOKEN)

# Use your actual channel ID
channel_id = -1001798059502  # Replace with your actual channel ID

# Define the timezone for Saudi Arabia
saudi_tz = pytz.timezone('Asia/Riyadh')

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

# Function to get the next word index from file
def get_next_word_index():
    try:
        with open('data/last_word_index.txt', 'r') as file:
            index = int(file.read().strip())
            logging.info(f"Loaded index: {index}")
            return index
    except FileNotFoundError:
        logging.error("last_word_index.txt not found, starting from index 0")
        return 0
    except ValueError:
        logging.error("Value error encountered, starting from index 0")
        return 0

# Function to update the last sent word index in file
def update_last_word_index(index):
    with open('data/last_word_index.txt', 'w') as file:
        file.write(str(index))
        logging.info(f"Updated index to: {index}")

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

# Initialize the scheduler
executors = {
    'default': ThreadPoolExecutor(1),
}
scheduler = BackgroundScheduler(executors=executors, timezone=pytz.utc)
scheduler.start()

# Function to schedule the job at 5:25 PM Saudi time
def schedule_job_at_saudi_time(hour, minute):
    # Get the next occurrence of the specified time in Saudi timezone
    now_saudi = datetime.now(saudi_tz)
    run_time_saudi = now_saudi.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if run_time_saudi <= now_saudi:
        run_time_saudi += timedelta(days=1)
    # Convert to UTC
    run_time_utc = run_time_saudi.astimezone(pytz.utc)

    # Schedule the job
    scheduler.add_job(send_word_of_the_day, 'cron', hour=run_time_utc.hour, minute=run_time_utc.minute)
    logging.info(f"Scheduled job at {run_time_utc.strftime('%Y-%m-%d %H:%M:%S %Z')} UTC (Corresponds to {hour}:{minute} Saudi Time)")

# Schedule the job at 11:25 AM Saudi time
schedule_job_at_saudi_time(11, 25)

# Start polling for bot commands
bot.polling(none_stop=True)
