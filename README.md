
# Aspen Plus Script Generator for Streamlit

This repository contains a Streamlit app that generates a Python automation script for Aspen Plus batch simulations.

## Included files
- `streamlit_app.py`: main Streamlit application.
- `requirements.txt`: dependencies required by the Streamlit app.

## What the app does
- Collects Aspen file path, input Excel path, output folder, Excel column names, and Aspen internal paths.
- Builds a ready-to-run Python automation script.
- Supports `.bkp`, `.apw`, and `.apwz` Aspen files.
- Includes a three-level convergence recovery strategy.
- Lets the user download the generated Python script directly from the browser.

## Deploy on Streamlit Community Cloud
1. Create a GitHub repository.
2. Upload `streamlit_app.py`, `requirements.txt`, and this `README.md`.
3. Go to Streamlit Community Cloud.
4. Create a new app from your GitHub repository.
5. Set the main file path to `streamlit_app.py`.
6. Deploy.

## Important note
The generated automation script is intended to run on a Windows machine with Aspen Plus installed and with Python packages such as `numpy`, `pandas`, `openpyxl`, `pywin32`, `tqdm`, and `reportlab` available.
Streamlit Cloud can host the generator app, but it cannot run Aspen Plus itself.
