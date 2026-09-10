#!/bin/zsh
# Double-click to launch the GWC Commissioner's Dashboard.
cd "$(dirname "$0")"
python3 -m streamlit run app.py
