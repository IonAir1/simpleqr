from simpleqr import IDMaker
import scrapetoexcel
import shutil


file_path = "sample/Dorm.DILIMAN.html"


def scrape():
    scrapetoexcel.html_to_excel(file_path, "to_input.xlsx", skip_pdf=False)#, space_before_pic=1) #add space before picture column
    inp = input("Proceed to QR Code Generation? y/enter=yes,r=retry,n=cancel \n")
    if inp.lower() == 'r':
        print("Retrying...")
        scrape()
    elif inp.lower() == 'n':
        print("Cacnceled!")
    elif inp.lower() in ['', 'y']:
        print("Proceeding...")
        IDMaker("idsettings.cfg").generate_ids()


scrape()