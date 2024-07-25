# The Office Script Translator

## Description

This Python script processes a CSV file containing scripts from the TV show "The Office" and translates each line into Turkish. It organizes the translated scripts into Word documents, sorted by season and episode.

## Data Source

The original CSV file containing "The Office" scripts was obtained from [Kaggle](https://www.kaggle.com/datasets/nasirkhalid24/the-office-us-complete-dialoguetranscript?resource=download).

## Features

- Reads CSV file with "The Office" scripts
- Translates each line from English to Turkish
- Organizes scripts into Word documents by season and episode
- Implements a retry mechanism for failed translations
- Tracks progress to allow for script interruption and resumption
- Batch translation for improved efficiency

## Prerequisites

- Python 3.6+

## Installation

1. Clone this repository or download the script.
2. Install the required packages:
   
   ```bash
   - pip install python-docx googletrans==3.1.0a0
## Usage

1. Ensure you have the CSV file 'The-Office-Lines-V4.csv' in the same directory as the script.
2. Run the script:
   ```bash
   -python script_name.py
3. The script will create a directory 'The_Office_Scripts' with subdirectories for each season, containing Word documents for each episode.

## Progress Tracking

The script saves progress in 'translation_progress.json'. If interrupted, it can be resumed from the last processed episode.
