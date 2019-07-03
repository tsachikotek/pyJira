import docx
from document import Document

print ('Documenting...')
document = Document()
document.add_heading('Document Title', 0)

print ('Saving...')
document.save('test.docx')