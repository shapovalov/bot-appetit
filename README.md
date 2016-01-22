Telegram Bot that displays lunch orders for Artec Moscow office.

Based on the Telegram Bot starter kit: https://github.com/yukuku/telebot

Reddit post: http://www.reddit.com/r/Telegram/comments/3b1pwl/create_your_own_telegram_bot_stepbystep/

Note that the Telegram bot token is not committed for security sake. Contact the maintainer to get access.

Usage
=====

To update the database (needs to be done each weak), perform the following actions.

1. Download the spreadsheet created by Masha as xlsx file.

2. Most of the work is done offline. Run the command to parse the file and form the order strings:
    ```
    ./dumpdb.py <spreadsheet>.xlsx db.p
    ```
    It dumps the orders to a pickle file.

3. Update the google app engine using its desktop app.

Bot Appétit.
