import pdfkit
import pypdf
from datetime import datetime as dt
import csv
import sys

class amazon_expense_gen():
    def __init__(self, url_txt = "amazon_urls.txt", who = "Carl", quantitiy = [], proj = "1601010642"):
        self.txt = url_txt
        self.urls = []
        self.names = []
        self.pdf_paths = []
        self.Quantity = quantitiy


        self.proj = proj
        self.item = {
            "Date":dt.now().strftime("%Y-%m-%d"),
            "Who": who,
            "Supplier": "Amazon",
            "Phone":"",
            "Contact": "",
            "ASIN": "",
            "Quantity": "",
            "Description": "",
            "Cost": "",
            "Total": "",
            "Shipping" :"",
            "Tax" :"",
            "Order Total":"",
            "Delivery":"",
            "Chemical":"",
            "Proj" : self.proj,
            "Comments":"",
            "Link" :""
        }

        self.items = []
        
    def get_urls(self):
        with open(self.txt, 'r') as file:
            for line in file:
                self.urls.append(line.strip())
        if self.Quantity == []:
            self.Quantity = ['1'] * len(self.urls)
            
    def generate_pdf_from_url(self):
        for url in self.urls:
            name = url.split('/')[3]
            self.names.append(name)
            pdf_path = f"{name}.pdf"
            self.pdf_paths.append(pdf_path)
            try:
                config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")
                pdfkit.from_url(url, pdf_path, configuration=config, verbose = True)
            except ConnectionRefusedError:
                continue

    def read_pdf(self):
        i = 0
        for pdf_path in self.pdf_paths:
            item = self.item.copy()
            reader = pypdf.PdfReader(pdf_path)
            item["Description"] = self.names[i]
            item["Quantity"] = self.Quantity[i]
            item["Link"] = self.urls[i]
            
            j = 0
            for page in reader.pages:
                
                if j == 0:#attempting to use in stock
                    first_page = page.extract_text()
                    t2 = first_page[first_page.index("In Stock"): first_page.index("In Stock") + 20]
                    t3 = t2.split("\n")
                    price = f"{t3[1]},{t3[3]}"
                    item["Cost"] = price
                    total = str(float(price.replace(",", ".")) * float(item["Quantity"])).replace(".", ",")
                    item["Total"] = f"{total}"
                    item["Order Total"] = f"{total}"
                try:
                    pt = page.extract_text() #page text
                    asin = pt[pt.index("ASIN") + 7:pt.index("ASIN") + 20]
                    asin = asin.split('\n')[0].strip()
                    item["ASIN"] = asin 
                        

                except ValueError:
                    item["ASIN"] = "Not found"
                j += 1
                
            self.items.append(item)
            i += 1
    
    def to_csv(self):
        name = f"amazon_expense_{dt.now().strftime('%Y-%m-%d')}.csv"
        csv_path = f"{name}"
        fns = ["Date", "Who", "Supplier", "Phone", "Contact", "ASIN", "Quantity", "Description", "Cost", "Total", "Shipping", "Tax", "Order Total", "Delivery", "Chemical", "Proj", "Comments", "Link"]
        
        with open(csv_path, 'a', newline='', encoding = "utf-8") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=fns)
            writer.writeheader()
        
            for item in self.items:
                #stripping items 
                for key,value in item.items():
                    value.replace("\u200b", " ")
                    item[key] = value.strip()

            writer.writerows(self.items)
#Date of Request Requested By Supplier Phone Contact Part #	Quantity Description Cost (each) Total Shipping Tax	Order Total	Tentative Ship/Delivery Date Chemical Prior Approval Status	Project #	Comments	Link
if __name__ == "__main__":
    print(sys.argv)
    if len(sys.argv) > 1:
        who = sys.argv[1]
        try:
            quantity = sys.argv[2].split(",")
        except IndexError:
            quantity = []
        temp = amazon_expense_gen(who=who, quantitiy=quantity)
    else:
        temp = amazon_expense_gen()
    
    temp.get_urls()
    temp.generate_pdf_from_url()
    temp.read_pdf()
    temp.to_csv()