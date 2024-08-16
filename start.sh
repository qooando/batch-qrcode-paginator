#!/bin/bash
#python3 -m venv  --system-site-packages ./venv
source ./venv/bin/activate
pip3 install -r requirements.txt
python3 ./make_qrcodes.py

