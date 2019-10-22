#!/bin/bash
#
cd ../
rm -r package
rm ../ultraInvoice.zip
pip install --target ./package xlutils xlrd xlwt requests
cd package
zip -r ../ultraInvoice.zip .
cd ..
zip -g ultraInvoice.zip *.py
aws lambda update-function-code --function-name UltraInvoice --zip-file fileb://ultraInvoice.zip


