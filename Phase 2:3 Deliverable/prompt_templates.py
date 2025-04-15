# prompt_templates.py

# ENRICH_PRESENTATION_PROMPT = """
# You are generating a presentation from a document.
# Return the slides as a **JSON array**.
# For each slide, include:

# {{
#     "title": "Slide title",
#     "text": "Bullet points or text content",
#     "image": "path/to/image.png",  // Select the most relevant image based on its context
#     "table": [["Row 1 Col 1", "Row 1 Col 2"], ["Row 2 Col 1", "Row 2 Col 2"]]  // optional
# }}

# Document text:
# {text_content}

# Available images with context (use to assign images meaningfully to slides):
# {image_paths}

# Available table data:
# {table_data}

# Guidelines:
# - **The number of slides should be determined by the document's topics and subtopics, not limited by the number of images.**
# - For each slide, focus on a distinct topic or subtopic from the document.
# - Generate multiple slides for major sections if needed, breaking down content into digestible slides.
# - For each slide, **if a relevant table is available, include it in the 'table' field in the response.**
# - Output must include the 'table' field for every slide, even if it is empty ("table": []).
# - Assign each image to only one slide. **Do not recycle images across multiple slides.**
# - If no suitable image is available for a slide, leave the "image" field empty.
# - Prioritize images whose context matches the slide topic or title.
# - Output must be a valid JSON array.
# - Keep your response clean, valid JSON, and do not add extra text.
# - Maintain bullet points in the 'text' field.
# - Example format:
# [
#   {{
#     "title": "Sample Slide",
#     "text": "• First point\\n• Second point",
#     "image": "./images/sample.png",
#     "table": []
#   }}
# ]
# Now generate the slide data.
# """


ENRICH_PRESENTATION_PROMPT = """
You are generating a presentation from a document.
Return the slides as a **JSON array**.
For each slide, include:

{{
    "title": "Slide title",
    "text": "Bullet points or text content",
    "image": "path/to/image.png",  // Select the most relevant image based on its context
    "table": [["Row 1 Col 1", "Row 1 Col 2"], ["Row 2 Col 1", "Row 2 Col 2"]]  // optional
}}

Document text:
{text_content}

Available images with context (use to assign images meaningfully to slides):
{image_paths}

Available table data:
{table_data}

Guidelines:
- **The number of slides should be determined by the document's topics and subtopics, not limited by the number of images.**
- For each slide, focus on a distinct topic or subtopic from the document.
- **Split long sections into multiple slides if they contain many bullet points, multiple concepts, or detailed breakdowns.**
- **Do not crowd a slide with too much text. Prefer multiple shorter slides over one long one.**
- Generate multiple slides for major sections if needed, breaking down content into digestible chunks.
- For each slide, **if a relevant table is available, include it in the 'table' field in the response.**
- Output must include the 'table' field for every slide, even if it is empty ("table": []).
- Assign each image to only one slide. **Do not recycle images across multiple slides.**
- If no suitable image is available for a slide, leave the "image" field empty.
- Prioritize images whose context matches the slide topic or title.
- Output must be a valid JSON array.
- Keep your response clean, valid JSON, and do not add extra text.
- Maintain bullet points in the 'text' field.
- Example format:
[
  {{
    "title": "Sample Slide",
    "text": "• First point\\n• Second point",
    "image": "./images/sample.png",
    "table": []
  }}
]
Now generate the slide data.
"""





# ENRICH_PRESENTATION_PROMPT = """
# You are generating a presentation from a document.
# Return the slides as a **JSON array**.
# For each slide, include:

# {{
#     "title": "Slide title",
#     "text": "Bullet points or text content",
#     "image": "path/to/image.png",  // Select the most relevant image based on its context
#     "table": [["Row 1 Col 1", "Row 1 Col 2"], ["Row 2 Col 1", "Row 2 Col 2"]]  // optional
# }}

# Document text:
# {text_content}

# Available images with context (use to assign images meaningfully to slides):
# {image_paths}

# Available table data:
# {table_data}

# Guidelines:
# - For each slide, **you must select the most contextually relevant image from the provided list.**
# - Do not leave the "image" field empty.
# - Even if no perfect match exists, assign the image with the most relevant context.
# - Prioritize images whose context matches the slide topic or title.
# - Ensure every slide has an "image" field with a valid path from the provided list.
# - Output must be a valid JSON array.
# - Keep your response clean, valid JSON, and do not add extra text.
# - Maintain bullet points in the 'text' field.
# - Example format:
# [
#   {{
#     "title": "Sample Slide",
#     "text": "• First point\\n• Second point",
#     "image": "./images/sample.png",
#     "table": []
#   }}
# ]
# Now generate the slide data.
# """






EXTRACT_TOPICS_MARKERS_TEMPLATE = """
Analyze the given document and extract key topics, following these guidelines:

1. Key Topic Identification:
   - Topics should represent major sections or themes in the document.
   - Each key topic should be substantial enough for at least one slide with 3-5 bullet points, potentially spanning multiple slides.
   - Topics should be broad enough to encompass multiple related points but specific enough to avoid overlap.
   - Identify topics in the order they appear in the document.
   - Consider a new topic when there's a clear shift in the main subject, signaled by transitional phrases, new headings, or a distinct change in content focus.
   - If a topic recurs, don't create a new entry unless it's substantially expanded upon.

2. Key Topic Documentation:
   - For each key topic, create a detailed name that sums up the idea of the section or theme it represents. 
   - Next, provide the first ten words of the section that the key topic represents.

3. Provide the output in the following format:
**<key topic 1>**
first ten words of the section or theme that the key topic 1 represents
**<key topic 2>**
first ten words of the section or theme that the key topic 2 represents

Document to analyze:
'''
{{content}}
'''

"""

GENERATE_SLIDE_CONTENT_TEMPLATE = """
You will be given a key topic, and a document portion, which provide detail about the key topic. Your task is to create slides based on the document portion. Follow these steps:

1. Identify the relevant section of the document between the given starting lines.
2. Analyze this section and create slides with titles and bullet points.

Guidelines:
- The number of slides can be 1-5, depending on the amount of **non-repetitive information** in the relevant section of the key topic.
- **Make sure the information on all the slides under same topic is non-repetitive.**
- Avoid generating multiple slides for the same topic.
- Present slides in the order that the information appears in the document.
- Each slide should have 4-6 concise bullet points, each containing a single key idea or fact.
- Use concise phrases or short sentences for bullet points, focusing on conveying key information clearly and succinctly.
- If information seems relevant to multiple topics, include it in the current topic's slides, as it appears first in the document.
- Avoid redundancy across slides within the same key topic.
- **Do not include additional commentary, explanations, or “Note:” sections. Provide only the slide titles and bullet points.**

Output Format:
**paste slide title here**
paste point 1 here
paste point 2 here
paste point 3 here

Inputs:
Key Topic: '''{{topic}}'''

Document portion:'''
{{contentSegment}}
'''

Please create slides based on the document portion, following the guidelines provided. Ensure that the slides comprehensively cover the key topic without unnecessary repetition.
**Output only the slide content** (titles + bullet points). **Do not add any extra notes or remarks. MAKE SURE THIS IS FOLLOWED.**.
"""