### emailer 
this exe is to check emails for a set period of time and then make a .csv file to compose everything into amt orders, currently only works for amazon

Email comp.
To use this system the email must be composed with the subject: amazon order then have a body as shown below

<url1>::number of url1 needed
<url2>::number of url2 needed

if not quantity is added it will be assumed to be 1

to run and check for emails over the last 24 hours just run the email.exe it can take custom times tho with the arg below
email.exe <start> <end>
times must beformatted as %m/%d/%Y-%H:%M an example is shown below

email.exe 08/10/2025-11:04 08/12/2025-11:04    

## explained 
This code defines an email_agent class that automates the process of extracting Amazon order information from Outlook emails received within a specified time range. The class uses the win32com.client library to interact with Microsoft Outlook, allowing it to search the Inbox for emails whose subject matches "amazon order". The time range for searching is set by default to the previous 24 hours, but can be customized via command-line arguments.

Within the find_orders method, the code connects to Outlook, retrieves the Inbox folder, and applies a filter to select emails received between the given start and end times. For each matching email, it checks if the subject is "amazon order" (case-insensitive), then parses the sender's name and extracts URLs from the email body. If a URL contains a "::" separator, it assumes the following value represents the quantity ordered; otherwise, it defaults the quantity to "1". All found URLs and quantities are stored in lists for later use.

The to_order_txt method writes the extracted Amazon order URLs and associated sender names to a text file called amazon_urls.txt, formatting each line as "name::url". This prepares the data for further processing.

In the main execution block, the script optionally accepts start and end times as command-line arguments, parses them, and creates an email_agent instance. It then calls the methods to find orders and write them to a file. Next, it instantiates an amazon_expense_gen object (from an imported module), passing the list of quantities, and runs several methods to process the URLs, generate PDFs, read PDF contents, and export the results to a CSV file. Finally, it attempts to clean up by deleting any generated PDF files.

Overall, the code automates the workflow of extracting Amazon order details from Outlook emails, saving them, and generating expense reports, with robust error handling and support for time-based filtering.

### usage of the amazon system
This exe is used to get a pdf from an amazon url then decode it into a csv file

Usage: 
1. open amazon_url.txt
2. add urls of all items you want to order, each url should be on its own line as show in the the sample below
    https://www.url1.com
    https://www.url2.com
    https://www.url3.com
3. cd into the exe dir
4. run amazon.exe <who> <quanity> see below for quanity argument parsing
    for example using the above url if I wanted one of each I leave quanity blank 
    if I want 2 of each the quanity argument is as follows 2,2,2

5. a new csv file should be made, open in excel and copy the cols besides the first over to the amt order excel sheet


example command steve wants to use the above urls to order to of each

amazon.exe Steve 2,2,2


git at https://github.com/cawilsonvitro/amazon_to_amt_orderf

## explained 
This code processes Amazon order information by extracting relevant details from PDF pages, likely representing order receipts or invoices. For each item in an order, it sets the item's description, quantity, and link using pre-populated lists (self.names, self.Quantity, and self.urls). It then iterates through each page of the PDF using a reader object.

On the first page (j == 0), the code attempts to extract price information by searching for the text "In Stock" (case-sensitive first, then lowercase if not found). It slices a portion of text around this keyword and splits it into lines, assuming that price details are located at specific line indices (t3[1] and t3[3]). These values are combined into a string, assigned to the item's "Cost", and used to calculate the total cost by multiplying with the item's quantity. The total is formatted and stored in both "Total" and "Order Total" fields.

The code also tries to extract the ASIN (Amazon Standard Identification Number) from each page. If the item's "ASIN" field is empty, it searches for the "ASIN" keyword in the page text, slices out a segment, splits it, and assigns the value found on the next line. If the keyword is not found, it sets the ASIN to an empty string.

A key detail is the reliance on specific text patterns and line positions, which may be brittle if the PDF format changes. The code uses exception handling to manage cases where expected keywords are missing, ensuring that the process continues without crashing. The loop variable j helps distinguish the first page from subsequent pages, as price information is only extracted from the first page. Overall, this snippet automates the extraction of price and ASIN details from Amazon order PDFs, populating a structured item dictionary for further processing.