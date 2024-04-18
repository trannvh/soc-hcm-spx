#!/usr/bin/env bash
sudo apt update --fix-missing
sudo apt install -y python3-pip
sudo apt install -y ./wkhtmltox_0.12.6.1-2.jammy_amd64.deb
pip install gspread
pip install pandas
pip install imgkit
pip install jinja2
chmod -R 777 ./../temp