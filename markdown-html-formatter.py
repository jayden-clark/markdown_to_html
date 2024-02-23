import re

from docx import Document


def replace_links_and_images(docx_path):
    doc = Document(docx_path)
    for paragraph in doc.paragraphs:
        markdown_links = re.findall(
            r"\[([^\]]+)\]\((http[s]?://[^\)]+)\)", paragraph.text
        )
        for link_text, url in markdown_links:
            new_link = f'<a href="{url}" style="color: blue; text-decoration: underline;">{link_text}</a>'
            paragraph.text = paragraph.text.replace(f"[{link_text}]({url})", new_link)

    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_name = rel.target_ref.split("/")[-1]  
            new_img_html = f'<div><img src="/images/blog/{img_name}" alt="" style="width:100%" /></div>'
            # Find the paragraph containing the image and replace it
            for paragraph in doc.paragraphs:
                if img_name in paragraph.text:
                    paragraph.text = new_img_html

    # Save and overwrite the modified document
    doc.save(docx_path)


replace_links_and_images(input("Enter .docx file name"))
