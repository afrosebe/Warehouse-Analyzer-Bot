#!/bin/bash
# Start Warehouse Auction Analyzer Bot
cd /data/data/com.termux/files/home/warehouse-analyzer
export BOT_TOKEN="8932041732:AAE3qq-jv3t9FpWngCZHTAn-Jgzbmcb_nn0"
nohup python3 bot.py > bot.log 2>&1 &
echo "Bot started! PID: $!"
echo "Check status: tail -f bot.log"