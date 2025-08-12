import win32com.client
import datetime as dt
from zoneinfo import ZoneInfo
from amazon import amazon_expense_gen


class email_agent():
    def __init__(self,
                 start_time = dt.datetime.now(ZoneInfo('America/New_York')) - dt.timedelta(days=1), 
                 end_time = dt.datetime.now(ZoneInfo('America/New_York'))):
        self.start_time = start_time
        self.end_time = end_time
        
        return
    
    def find_orders(self):
               
        try:
            # Dispatch Outlook Application
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            # Get the Inbox folder (folder index 6 is typically the Inbox)
            inbox = namespace.GetDefaultFolder(6)
                # Example: Emails received in the last 24 hours
           

            start_time_str = self.start_time.strftime('%m/%d/%Y %H:%M ')
            end_time_str = self.end_time.strftime('%m/%d/%Y %H:%M ')
            print(f"Fetching emails from {start_time_str} to {end_time_str}")
            filter_string = f"[ReceivedTime] >= '{start_time_str}' AND [ReceivedTime] <= '{end_time_str}'"
            messages = inbox.Items.Restrict(filter_string)
            amazon_orders = []
            for message in messages:
            #    print(f"Subject: {message.Subject}, Sender: {message.SenderName}, Received: {message.ReceivedTime}")
                if "amazon order" == message.Subject.lower():
                    print(f"Amazon Order Found: {message.Subject}")
                    amazon_orders.append(message.Body)

        except Exception as e:
            pass