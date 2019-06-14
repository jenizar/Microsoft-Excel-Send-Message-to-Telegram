# Microsoft-Excel-Send-Message-to-Telegram
Microsoft Excel Send Message to Telegram

![alt text](https://github.com/jenizar/Microsoft-Excel-Send-Message-to-Telegram/blob/master/images/UserInterfaceGetResponse.PNG)

Step by step :
1. Create Telegram Bot
2. Get Telegram Token API
3. Start Conversations in your new Telegram Bot (type more messages) 
4. Get chat_id (https://api.telegram.org/bot<token>/getUpdates)
e.g. https://api.telegram.org/bot446024890.../getUpdates
5. Get info your Bot (https://api.telegram.org/bot<token>/getMe)
e.g. https://api.telegram.org/bot446024890.../getMe
6. Create VBA Excel Program
7. Run VBA Excel Program
filled parameter : chat_id & Message
8. Enjoy

Source code:
https://github.com/jenizar/Microsoft-Excel-Send-Message-to-Telegram

References:
Telegram Bot API
1. Cara Membuat Bot Telegram - Blajar Bot Telegram #1 
https://youtu.be/pgDabVwUwLc
2. Cara Menggunakan Telegram Bot API - Blajar Bot Telegram #2
https://youtu.be/3IQKb_orBlo
3. Telegram Bot API
https://core.telegram.org/bots/api

VBA Microsoft Excel
1. https://stackoverflow.com/questions/1820345/perform-http-post-from-within-excel-and-parse-results

2. https://stackoverflow.com/questions/158633/how-can-i-send-an-http-post-request-to-a-server-from-excel-using-vba
