import os
from docx import Document
from dotenv import load_dotenv

#l load .env variables
load_dotenv()

# settings
input_dir = os.getenv("INPUT_DIR")
output_dir = os.getenv("OUTPUT_DIR")
find_text = os.getenv("FIND_TEXT")
replace_text = os.getenv("REPLACE_TEXT")

# check for output directory
os.makedirs(output_dir, exist_ok=True)

def replace_in_docx(input_path, output_path, find, replace):
	doc = Document(input_path)
	for p in doc.paragraphs:
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