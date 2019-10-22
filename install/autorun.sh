#!/bin/sh
#
echo "Starting MonthlyReport.py path=/home/ec2-user..."
python /home/ec2-user/ultraInvoice/MonthlyReport.py path=/home/ec2-user

echo "MonthlyReport.py Finished. and shutdown after 3 minutes"
sudo shutdown -h +3 "3 minutes after shutdown"
