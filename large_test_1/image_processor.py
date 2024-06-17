from xlsx2html import xlsx2html
import pdfkit
import pdf2image

xlsx2html('large_test_1.xlsx', 'large_test_1.html')

pdfkit.from_file('large_test_1.html', 'large_test_1.pdf')

images = pdf2image.convert_from_path('large_test_1.pdf')

for i in range(len(images)):
    images[i].save(f"large_test_1_page{i}.jpeg", 'jpeg')
