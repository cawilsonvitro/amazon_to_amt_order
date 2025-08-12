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


git at https://github.com/cawilsonvitro/amazon_to_amt_order