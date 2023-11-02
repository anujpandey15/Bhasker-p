git clone <https://github.com/anujpandey15/Bhasker-p/tree/main>
pip install python-pptx
from pptx import Presentation
import openpyxl
# Load the PowerPoint file
presentation = Presentation('https://docs.google.com/presentation/d/1TVhYe_UCTGwCdfhjF75X-N1PieSFuY6o/edit?usp=drivesdk&ouid=116581856421236729277&rtpof=true&sd=true')
# Initialize an Excel workbook
wb = openpyxl.Workbook()
ws = wb.active
# Iterate through slides and extract text
for slide in presentation.slides:
for shape in slide.shapes:
if hasattr(shape, "text"):
ws.append([shape.text])
# Save the Excel file
wb.save('output.xlsx')





































