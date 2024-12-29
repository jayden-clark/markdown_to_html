import glob
import os
import re

from docx import Document

# To use, convert the docx to markdown formatting first using https://word2md.com/
def replace_links_and_images(docx_path):
    doc = Document(docx_path)
    for paragraph in doc.paragraphs:
        markdown_links = re.findall(
            r"\[([^\]]+)\]\((http[s]?://[^\)]+)\)", paragraph.text
        )
        for link_text, url in markdown_links:
            new_link = f'<a href="{url}" style="color: blue; text-decoration: underline;">{link_text}</a>'
            paragraph.text = paragraph.text.replace(f"[{link_text}]({url})", new_link)

        markdown_images = re.findall(
            r"!\[\]\((data:/images/blog/[^)]+)\)", paragraph.text
        )
        for img_path in markdown_images:
            img_name = img_path.split("/")[-1]
            new_img_html = f'<div><img src="/images/blog/{img_name}" alt="" style="width:100%" /></div>'
            paragraph.text = paragraph.text.replace(f"![]({img_path})", new_img_html)

    # Overwrite original file
    doc.save(docx_path)


directory_path = "docs"

# Loop through all .docx files in the directory
for docx_file in glob.glob(os.path.join(directory_path, "*.docx")):
    print(f"Processing {docx_file}...")
    replace_links_and_images(docx_file)
