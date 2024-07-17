# ChatGPT Email Parser

This is a Python script that uses OpenAI's GPT-4 model to parse and filter emails. It can read emails directly from a specified mailbox or from a file. It also provides the functionality to filter emails by a specific value.

## Installation

Before running the script, you need to install the required dependencies. You can do this by running the following command in your terminal:


> activate your virtual environment using the command below

```bash 
source venv/bin/activate

pip install -r requirements.txt
```
or  

> run the script below to activate your virtual environment

```bash
./activate_venv.sh
```
```bash
Please make sure you have Python 3.6 or above installed on your system.

## Usage

You can run the script from the command line using the `python` command followed by the script name `chatgpt.py` and any arguments you wish to pass.

Here are the available arguments:

- `--mailbox`: The mailbox to read the emails from.
- `--filter_by`: The field to filter the emails by. This could be 'subject', 'sender', etc.
- `--filter_by_value`: The value to filter the emails by.
- `--is_file`: A boolean value indicating whether to read the email from a file.
- `--source_file_name`: The name of the file to read the email from.
- `--start_date`: The start date to filter the emails by.
- `--end_date`: The end date to filter the emails by.
- `--num_emails`: The number of emails to read.

Here are some examples of how to run the script:

```bash

python chatgpt.py --mailbox inbox --filter_by subject --filter_by_value "Second content of upwork"

or 

python chatgpt.py --mailbox inbox --filter_by "subject" --filter_by_value "Second content of upwork"

or 

python chatgpt.py --mailbox inbox --filter_by "from" --filter_by_value "upwork"

or 

python chatgpt.py --mailbox inbox --filter_by "to" --filter_by_value "upwork"

or 

python chatgpt.py --mailbox inbox  --start_date "20-06-2023" --end_date "22-06-2023" --num_emails 10

or 

python chatgpt.py --is_file True --source_file_name "email.txt"
```

Please replace the argument values with those that suit your needs.

# You can generate your app password from [APP PASSWORD](https://myaccount.google.com/u/0/apppasswords) and use it in the script to access your gmail account. 
