import win32com.client
import datetime as dt
from zoneinfo import ZoneInfo
from amazon import amazon_expense_gen
import os
import sys
class email_agent():
    def __init__(self,
                 start_time = dt.datetime.now(ZoneInfo('America/New_York')) - dt.timedelta(days=1), 
                 end_time = dt.datetime.now(ZoneInfo('America/New_York'))):
        self.start_time = start_time
        self.end_time = end_time
        self.amazon_orders = []
        self.names = []
        self.quantity = []  
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
            for message in messages:
            #    print(f"Subject: {message.Subject}, Sender: {message.SenderName}, Received: {message.ReceivedTime}")
                if "amazon order" == message.Subject.lower():
                    self.names.append(message.SenderName.split(",")[-1])
                    print(f"Amazon Order Found: {message.SenderName}")
                    urls = message.Body.split("https://")[1:]
                    #getting quantity out
                    k = 0
                    for url in urls:
                        if "::" in url:
                            quantity = url.split("::")
                            quantity = [str(i).strip() for i in quantity]
                            self.quantity.append(quantity[1])
                      
                        else:
                            self.quantity.append("1")
                        k += 1
                    full_urls = []
                    for url in urls:
                        full_url = "https://" + url.split()[0]
                        full_urls.append(full_url)
                        
                    self.amazon_orders.append(full_urls)
        except Exception as e:
            pass
        
    def to_order_txt(self):
        with open("amazon_urls.txt", "w") as f:
            i = 0
            for order in self.amazon_orders:
                who = self.names[i] 
                for url in order:
                    f.write(who +"::"+url.strip() + "\n")
                i += 1 
if __name__ == "__main__":
    format_string = "%m/%d/%Y-%H:%M"
    if len(sys.argv) > 1:
        try:
            start = dt.datetime.strptime(sys.argv[1], format_string)
            end = dt.datetime.strptime(sys.argv[2], format_string)
            agent = email_agent(start_time=start, end_time=end)
        except Exception as e:
            print(e)
            agent = email_agent()
    else:
        agent = email_agent()
    agent.find_orders()
    agent.to_order_txt()
    report_gen = amazon_expense_gen(quantity=agent.quantity)
    report_gen.get_urls()
    report_gen.generate_pdf_from_url()
    report_gen.read_pdf()
    report_gen.to_csv()
    
    # #cleaning up
    for pdf in report_gen.pdf_paths:
        print(pdf)
        try:
            os.remove(pdf)
        except Exception as e:
            print(e)
            pass
#how to deal with  str args
# start = "08/12/2025 11:04"
# format_string = "%m/%d/%Y-%H:%M"

# datetime_obj = dt.datetime.strptime(start, format_string)
# print(datetime_obj)