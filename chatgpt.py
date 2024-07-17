import datetime
import email.message
import re
import json, time, openai, imaplib, email, pandas as pd, os , argparse
from dotenv import dotenv_values

class ParseEmail:
    def __init__(self, mailbox_selected = "inbox", filter_by = None, filter_by_value = None, is_file = False, source_file_name = "" ,start_date = None, end_date = None, num_emails = None):
        """
        Initialize the ParseEmail class

        Args:
            mailbox_selected (str, optional): Selected mailbox. Defaults to "inbox".
            filter_by (str, optional): Filter the email by. Defaults to None.
            filter_by_value (str, optional): Filter the email by value. Defaults to None.
            is_file (bool, optional): Check if the source is a file. Defaults to False.
            source_file_name (str, optional): Source file name. Defaults to "".
            start_date (str, optional): Start date to retrieve emails from. Defaults to None.
            end_date (str, optional): End date to retrieve emails from. Defaults to None.
            num_emails (int, optional): Number of emails to retrieve. Defaults to None.
        
        Raises:
            FileNotFoundError: File not found
            ValueError: Invalid file format. Only .txt files are supported
            ValueError: Invalid filter_by parameter
        """
        # Get the current directory
        self.current_dir = os.path.dirname(os.path.realpath(__file__))
        # Create the directories to store the email and the result
        os.makedirs(os.path.join(self.current_dir, "RESULT"), exist_ok=True)
        # Create the directories to get the source file
        os.makedirs(os.path.join(self.current_dir, "SOURCE"), exist_ok=True)
        # Set the excel file name
        self.excel_file_name = os.path.join(self.current_dir, "RESULT","final_parsed.xlsx")
        
        # Get the credentials from the .env file
        env = dotenv_values("credentials/.env")
        self.all_emails = []
        self.query = None
        # Set the API key
        self.api_key = env["OPENAI_API_KEY"]
        # Initialize the OpenAI API
        self.openai = openai
        # Set the API key
        self.openai.api_key = self.api_key
        self.login__result = False
        if is_file:
            self.file_name = source_file_name
            self.email_subject = self.file_name.split(".")[0] + " " + "OpenAI ChatGPT"
            # Set the file path
            self.file_path = os.path.join(self.current_dir,"SOURCE", self.file_name)
            
            # Set the excel file name
            self.excel_file_name = os.path.join(self.current_dir, "RESULT", self.file_name.split(".")[0] + ".xlsx")
            
            # Check if the file exists
            if not os.path.exists(self.file_path):
                raise FileNotFoundError("File not found")
    
            # Check if the file format is .txt
            if not self.file_name.endswith(".txt"):
                raise ValueError("Invalid file format. Only .txt files are supported")
            
            # Read the email from the file            
            with open(self.file_path, "r") as f:
                self.email_body = f.read()
            self.openi_ask_format()
            
        elif not is_file and filter_by is not None and filter_by_value is not None:   
            # Set the credentials
            self.my_email = env["MY_EMAIL"]
            self.app_password = env["APP_PASSWORD"]
            self.filters = {"subject": "SUBJECT", "from": "FROM", "to": "TO", "body": "BODY"}
    
            # Check if the filter_by parameter is valid
            if self.filters.get(filter_by.lower()) is None:
                # If not, raise an exception
                raise ValueError("Invalid filter_by parameter")
            
            # Set the query to filter the email
            self.query = f'{self.filters[filter_by]} "{filter_by_value}"'
            # Set the mailbox
            self.mailbox_list = {"inbox": "INBOX", "sent": "[Gmail]/Sent Mail", "drafts": "[Gmail]/Drafts", "archive": "[Gmail]/All Mail", "spam": "[Gmail]/Spam", "trash": "[Gmail]/Trash"}
            # Set the selected mailbox
            self.selected_mailbox = self.mailbox_list[mailbox_selected]
            # Get the email body
            self.login__result = self.login_()
        else:
            self.my_email = env["MY_EMAIL"]
            self.app_password = env["APP_PASSWORD"]
            self.selected_mailbox = "INBOX"
            self.start_date = start_date
            self.end_date = end_date
            self.num_emails = num_emails
            self.login__result = self.login_()
            
        
        
    
    def login_(self):
        """
        Login to the email using the credentials provided

        Returns:
            Exception: Exception if an error occurs
        """
        while True:
            try:
                # Connect to the mailbox
                self.mailbox = imaplib.IMAP4_SSL("imap.gmail.com")
                # Login to the mailbox using the credentials
                self.mailbox.login(self.my_email, self.app_password)
                if self.mailbox.state == "AUTH":
                    print("\033[92mSuccessfully logged in fetching email ...\033[0m")
                    return True
            # If an exception occurs, print the exception and retry the process
            except Exception as e:
                time.sleep(1)
                if "Web login required" in str(e):
                    #print "Please enable access for less secure apps in your Gmail account" in red color
                    print("\033[91mPlease enable access for less secure apps in your Gmail account\033[0m")
                    return False
                elif "Invalid credentials" in str(e):
                    print("\033[91mInvalid credentials. Please check the credentials in the /credentials/.env file\033[0m")
                    return False
                else:
                    # in yellow color print "An error occurred. Please try again" and the exception
                    print("\033[93mAn error occurred. Please try again\033[0m", e.__str__())
                    time.sleep(5)
      
     
    def get_email(self):
        """
        Get the email from Gmail using filter by and filter by value query
        """
        # Select the mailbox
        self.mailbox.select(self.selected_mailbox)
        # Get the email ids of the emails in the mailbox
        result, data = self.mailbox.search(None, f'{self.query}')
        # Get the the latest email id
        email_id = data[0].split()[-1]
        # Fetch the email with the specified email id
        result, data = self.mailbox.fetch(email_id, "(RFC822)")
        # Get the raw email data
        raw_email = data[0][1]
        # Get the email message
        email_message = email.message_from_bytes(raw_email)
        # Get the email body
        self.get_email_body(email_message)
        
        
    def retrieve_emails(self):
        """
        Retrieve emails from the mailbox

        Args:
            num_emails (int, optional): Number of emails to retrieve. Defaults to None.
            start_date (str, optional): Start date to retrieve emails from. Defaults to None.
            end_date (str, optional): End date to retrieve emails from. Defaults to None.

        Returns:
            list: List of emails
        """
        # Select the mailbox
        self.mailbox.select(self.selected_mailbox)
        # Get the email ids of the emails in the mailbox between the specified dates use SINCE
        search_start_date = datetime.datetime.strptime(self.start_date, "%d-%m-%Y").strftime("%d-%b-%Y")
        search_end_date = datetime.datetime.strptime(self.end_date, "%d-%m-%Y").strftime("%d-%b-%Y")
        result, data = self.mailbox.search(None, f'(SINCE "{search_start_date}" BEFORE "{search_end_date}")')
        # Get the email ids
        email_ids = data[0].split()
        # Iterate through the email ids
        for email_id in email_ids:
            
            # Fetch the email with the specified email id
            result, data = self.mailbox.fetch(email_id, "(RFC822)")
            # Get the raw email data
            raw_email = data[0][1]
            # Get the email message from the raw email data
            email_message = email.message_from_bytes(raw_email)
            try:
                email_date = datetime.datetime.strptime(email_message["Date"], "%a, %d %b %Y %H:%M:%S %z")
            except:
                try:
                    email_date = datetime.datetime.strptime(email_message["Date"], "%a, %d %b %Y %H:%M:%S %z (%Z)")
                except:
                    
                    try:
                        #email_date = Tue, 20 Jun 2023 03:15:04 -0500 (CDT) grab this Tue, 20 Jun 2023 only
                        email_date = datetime.datetime.strptime(email_message["Date"][0:16], "%a, %d %b %Y %H:%M")
                    except:
                        continue
            # Remove the timezone information
            email_date = email_date.replace(tzinfo=None)
            
            if email_message:
                # Get the email body using the email message
                self.get_email_body(email_message)
                
            # If the number of emails is specified and the number of emails is equal to the specified number of emails, break the loop   
            if self.num_emails is not None and len(self.all_emails) == self.num_emails:
                break
        return self.all_emails
    
    def get_email_body(self, email_message):
        """
        Get the email body from the email message and append it to the list of emails

        Args:
            email_message (email.message.Message): Email message
        """
        self.email_body = ""
        self.email_subject = email_message["Subject"]
        self.email_date = email_message["Date"]
        self.email_from = email_message["From"]
        self.email_to = email_message["To"]
        if isinstance(email_message, email.message.Message):
            if email_message.is_multipart():
                # Iterate through the parts of the email
                for part in email_message.walk():
                    # Get the content type of the part
                    content_type = part.get_content_type()
                    # Get the content disposition of the part
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # Get the body of the part and decode it
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    # Check if the content type is text/plain and the part is not an attachment
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        email_body = body
                        self.all_emails.append({"email_body": email_body, "email_subject": self.email_subject, "email_date": self.email_date, "email_from": self.email_from, "email_to": self.email_to})
            else:
                # Get the content type of the email message
                content_type = email_message.get_content_type()
                # Get the body of the email message and decode it
                body = email_message.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    email_body = body
                    self.all_emails.append({"email_body": email_body, "email_subject": self.email_subject, "email_date": self.email_date, "email_from": self.email_from, "email_to": self.email_to})
             
    def run(self):
        """
        Run the ChatGPT model to parse the email body into a list of key value pairs
        
        """
        if self.query is not None:
            self.get_email()
        else:
            # Retrieve the emails from the mailbox
            self.retrieve_emails()
        if self.all_emails:
            for email in self.all_emails:
                self.email_body = email["email_body"]
                self.email_subject = email["email_subject"]
                self.email_date = email["email_date"]
                self.email_from = email["email_from"]
                self.openi_ask_format()
                
        print("Email parsed successfully")
    
    def openi_ask_format(self):
        success = False
        while not success:
            # Get the response from the OpenAI API
            response = self.openai.chat.completions.create(
                    model="gpt-3.5-turbo-16k-0613",
                    messages=[
            {
            "role": "system",
            "content": "You will be provided with unstructured email data, and your task is to answer if you can parse it into a list of key value pairs each row should have same keys, That can be saved to an excel file.? say either 'YES, I CAN.' or NO, 'I CAN'T.' return only these dont include any other text"
            },
            {
            "role": "user",
            "content": self.email_body
            }
        ],
                )
            summary = response.choices[0].message.content
            try:
                summary = summary.upper()
            except:
                return
           
            if summary:
                success = True
                if "YES" in summary:
                    return self.openai_chat_response()
                else:
                    pass
                
            else:
                success = False
                
    def openai_chat_response(self):
    
        while True:
            # Get the response from the OpenAI API
            response = self.openai.chat.completions.create(
                model="gpt-3.5-turbo-16k-0613",
                messages=[
        {
        "role": "system",
        "content": "You will be provided with unstructured email data, and your task is to parse it into a list of key value pairs each row should have same keys, don't include any other text other than the list json key value pair, That can be saved to an excel file. make sure to prevent errors like 'file Expecting property name enclosed in double quotes and  Unterminated string starting at: line 1 column 2 (char 1)'. If you can do this task, please provide the parsed data in the format of a list of key value pairs."
        },
        {
        "role": "user",
        "content": self.email_body
        }
    ],
            )
            # Get the summary of the email from the response
            summary = response.choices[0].message.content
            if not summary:
                print("No response from the OpenAI API. Please try again")
                return
            else:
                try:
                    # Save the summary to an excel file
                    result = self.save_to_excel(summary)
                    if result:
                        return summary
                    else:
                        print("An error occurred. while saving the data to the excel file retrying ...")
                        continue
                except Exception as e:
                    continue
                
    def save_to_excel(self, list_of_dicts):
        """
        Save the parsed data to an excel file

        Args:
            list_of_dicts (str): list of key value pairs to save to the excel file
        """
        email_subject = re.sub(r"[^a-zA-Z0-9]+", ' ', self.email_subject).strip()
        self.excel_file_name = os.path.join(self.current_dir, "RESULT", email_subject + ".xlsx")
        
        try:
            
            # Convert the list of key value pairs to a dataframe
            df = pd.DataFrame(json.loads(list_of_dicts))
            # Save the dataframe to an excel file
            df.to_excel(self.excel_file_name, index=False)
            print("Data saved to excel file", self.excel_file_name)
            return True
        except Exception as e:
            try:
                if "If using all scalar values" in list_of_dicts:
                    df = pd.DataFrame(list_of_dicts, index=[0])
                    df.to_excel(self.excel_file_name, index=False)
                    print("Data saved to excel file", self.excel_file_name)
                    return True
                else:
                    return False
            except:
                return False
                    
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="ChatGPT")
    parser.add_argument("--mailbox", type=str, default="inbox", help="Selected mailbox")
    parser.add_argument("--filter_by", type=str, default=None, help="Filter the email by")
    parser.add_argument("--filter_by_value", type=str, default=None, help="Filter the email by value")
    parser.add_argument("--is_file", type=bool, default=False, help="Check if the source is a file")
    parser.add_argument("--source_file_name", type=str, default="", help="Source file name")
    parser.add_argument("--start_date", type=str, default="20-06-2023", help="Start date to retrieve emails from")
    parser.add_argument("--end_date", type=str, default="22-06-2023", help="End date to retrieve emails from")
    parser.add_argument("--num_emails", type=int, default=10, help="Number of emails to retrieve")
    args = parser.parse_args()
    # Initialize the ParseEmail class
    parse_email = ParseEmail(args.mailbox, args.filter_by, args.filter_by_value, args.is_file, args.source_file_name, args.start_date, args.end_date, args.num_emails)
    # Run the ParseEmail class
    if parse_email.login__result:
        parse_email.run()
    
    # how to run the code
    
    # python chatgpt.py --mailbox inbox --filter_by "subject" --filter_by_value "Second content of upwork"
    # python chatgpt.py --mailbox inbox --filter_by "from" --filter_by_value "upwork"
    # python chatgpt.py --mailbox inbox --filter_by "to" --filter_by_value "upwork"
    # python chatgpt.py --mailbox inbox --end_date "22-06-2023" --start_date "20-06-2023" --num_emails 10
    # python chatgpt.py --is_file True --source_file_name "email.txt"
    