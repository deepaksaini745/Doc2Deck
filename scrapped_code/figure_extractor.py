# figure_extractor.py

import os
import re
import uuid
from docx import Document
from lxml import etree  # You need lxml if you want to climb the XML tree
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def extract_figures_from_docx(docx_path, output_folder="extracted_figures"):
    doc = Document(docx_path)
    os.makedirs(output_folder, exist_ok=True)

    figures_info = []
    figure_counter = 1

    for para in doc.paragraphs:
        paragraph_text = para.text.strip()
        for run in para.runs:
            drawing = run._element.xpath('.//w:drawing')
            if not drawing:
                continue
            for draw in drawing:
                blip = draw.xpath('.//a:blip')
                if not blip:
                    continue
                embed_rid = blip[0].get(qn('r:embed'))
                image_part = doc.part.related_parts[embed_rid]
                image_bytes = image_part.blob

                image_filename = f"figure_{figure_counter}_{uuid.uuid4().hex[:8]}.png"
                image_path = os.path.join(output_folder, image_filename)
                with open(image_path, 'wb') as f:
                    f.write(image_bytes)

                figures_info.append({
                    'id': figure_counter,
                    'type': 'image',
                    'path': image_path,
                    'paragraph_text': paragraph_text
                })
                figure_counter += 1

    return figures_info


def decide_slide_mapping(figures_info, text_slides):
    figure_to_slide_map = []

    for fig in figures_info:
        is_large = (fig['type'] == 'table')

        best_slide_index = None
        best_slide_title = None
        best_match_score = 0

        fig_text = fig.get('paragraph_text', "").strip()

        if fig_text:
            fig_words = set(re.split(r'\W+', fig_text.lower()))

            for i, slide in enumerate(text_slides):
                slide_text = (slide.get('title', "") + " " + slide.get('content', "")).lower()
                slide_words = set(re.split(r'\W+', slide_text))

                overlap = len(fig_words.intersection(slide_words))

                if overlap > best_match_score:
                    best_match_score = overlap
                    best_slide_index = i
                    best_slide_title = slide.get('title')

        if best_slide_title and not is_large:
            figure_to_slide_map.append({
                'figure_id': fig['id'],
                'figure_type': fig['type'],
                'figure_path': fig['path'] if fig['type'] == 'image' else None,
                'figure_slide_title': best_slide_title,
                'own_slide': False
            })
        else:
            figure_to_slide_map.append({
                'figure_id': fig['id'],
                'figure_type': fig['type'],
                'figure_path': fig['path'] if fig['type'] == 'image' else None,
                'figure_slide_title': None,
                'own_slide': True
            })

    return figure_to_slide_map



if __name__ == "__main__":
    """
    Example usage:
    
    1. Extract figures from a sample docx
    2. Suppose we have some text slides from your existing text extraction step
    3. Get a mapping for where each figure should go
    """
    sample_docx = "example.docx"  # Replace with your .docx file
    figures = extract_figures_from_docx(sample_docx, output_folder="temp_figures")
    # figures = extract_all_images(sample_docx, output_folder="temp_figures")

    # Example set of slides from your text-extraction logic
    text_slides_example = [
        {'title': 'Introduction', 'content': 'This presentation covers the project overview and objectives.'},
        {'title': 'Methodology', 'content': 'We used a step-by-step process to analyze the data and produce results.'},
        {'title': 'Results', 'content': 'The final results are shown in the chart below.'}
    ]

    mapping = decide_slide_mapping(figures, text_slides_example)

    print("Extracted Figures:")
    for f in figures:
        print(f"  {f}")

    print("\nFigure-to-Slide Mapping:")
    for m in mapping:
        print(m)
