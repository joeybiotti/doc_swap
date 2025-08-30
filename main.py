import os
from docx import Document

# settings
input_dir = 'input_dir'
output_dir = 'output_dir'
find_text = 'old phrase'
replace_text = 'new phrase'

# check for output directory
os.makedirs(output_dir, exist_ok=True)

def replace_in_docx(input_path, output_path, find, replace):
	doc = Document(input_path)
	for p in doc.paragraph:
		if find in p.text:
			inline = p.runs 
			for i in inline:
				if find in i.text:
					i.text = i.text.replace(find, replace)
	doc.save(output_path)

# loop through files
for file in os.listdir(input_dir):
	if file.endswith(".docx"):
		input_path = os.path.join(input_dir, file)
		output_path = os.path.join(output_dir,file)
		replace_in_docx(input_path, output_path, find_text, replace_text)
		print(f"Processed {file}")