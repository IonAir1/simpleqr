import scrapetoexcel

file_path = "sample/Dorm.DILIMAN.html"
output_path = "to_input.xlsx"

scrapetoexcel.html_to_excel(file_path, output_path, skip_pdf=False, sort_by='name')#, space_before_pic=1) #add space before picture column