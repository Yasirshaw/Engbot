import pandas as pd
import telebot
import schedule
import time

# Initialize the Telegram bot with your token
API_TOKEN = '7364384207:AAGFQEJ6-SdGoZSuSc3WATxwZjDnBlgFHj8'
bot = telebot.TeleBot(API_TOKEN)

# Use your actual Telegram user ID for private messages
your_user_id = 724083023  # Yasir's user ID

# Function to load the words and definitions from the Excel sheet
def load_words_from_excel():
    df = pd.read_excel('data/words.xlsx')  # Ensure your Excel file is in the right location
    return df

# Function to send the word of the day from the Excel sheet
def send_word_of_the_day():
    words_df = load_words_from_excel()

    # Get the first row (Word of the Day)
    word = words_df.iloc[0]['Word']
    arabic_def = words_df.iloc[0]['Arabic Definition']
    example_sentence = words_df.iloc[0]['Example Sentence']

    # Craft the message with improved spacing and Markdown formatting
    message = (
        f"🎓 *Word of the Day* | *كلمة اليوم*\n\n"
        f"**{word}** - **{arabic_def}**\n\n"
        f"📖 *Example | مثال*: \n\n"
        f"{example_sentence}\n\n"
        f"💡 *Try using the word in a sentence today!*\n"
        f"💡 *حاول استخدام الكلمة في جملة اليوم!* 😊"
    )

    # Send the message to your private chat
    bot.send_message(chat_id=your_user_id, text=message, parse_mode='Markdown')

# Function to manually trigger the word of the day via command
@bot.message_handler(commands=['send_word'])
def manual_word_of_the_day(message):
    send_word_of_the_day()

# Schedule the bot to send the word of the day at 1 PM every day
schedule.every().day.at("13:00").do(send_word_of_the_day)

# Polling and scheduling loop
while True:
    # Check for scheduled tasks (like sending at 1 PM)
    schedule.run_pending()

    # Polling for manual command triggers
    bot.polling(none_stop=True)

    # Prevent CPU overload
    time.sleep(60)  # Check every minute for scheduled tasks
