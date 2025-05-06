# prompt_templates.py

# ENRICH_PRESENTATION_PROMPT = """
# You are generating a presentation from a document.
# Return the slides as a JSON array.
# Each slide should include:

# {{
#   "title": "Slide title",
#   "text": "• bullet 1\\n• bullet 2",
#   "image": "path/to/image.png",
#   "table": [["Header 1", "Header 2"], ["Row 1 Col 1", "Row 1 Col 2"]]
# }}

# Document text:
# {text_content}

# Available images with context:
# {image_paths}

# Available table data:
# {table_data}

# # Guidelines:
# # - Each slide should have 4-6 concise bullet points, each containing a single key idea or fact.
# # - **The number of slides should be determined by the document's topics and subtopics, not limited by the number of images.**
# # - For each slide, focus on a distinct topic or subtopic from the document. 
# # - **Split long sections into multiple slides if they contain many bullet points, multiple concepts, or detailed breakdowns.**
# # - Split content into as many distinct slides as needed based on content structure. Do not limit to 8–10 slides.
# # - When splitting content into multiple slides, avoid generic titles like "Part 1", "Part 2", or "continued" unless no better option is possible.
# # - Instead, try to generate unique, descriptive slide titles that clearly reflect the key idea, subtopic, or focus of the slide.
# # - **Do not crowd a slide with too much text. Prefer multiple shorter slides over one long one.**
# # - **Also don't oversimplify slides with just 2-3 word bullets where explanation could be required.**
# # - Generate multiple slides for major sections if needed, breaking down content into digestible chunks.
# # - For each slide, **if a relevant table is available, include it in the 'table' field in the response.**
# # - Output must include the 'table' field for every slide, even if it is empty ("table": []).
# # - If you include a table on a slide, do not repeat the same information in bullet points. Choose one format per piece of content.
# # - Assign each image to only one slide. **Do not recycle images across multiple slides.**
# # - If no suitable image is available for a slide, leave the "image" field empty.
# # - Prioritize images whose context matches the slide topic or title.
# # - Output must be a valid JSON array.
# # - Keep your response clean, valid JSON, and do not add extra text.
# # - Maintain bullet points in the 'text' field.
# # - Example format:
# # [
# #   {{
# #     "title": "Sample Slide",
# #     "text": "• First point\\n• Second point",
# #     "image": "./images/sample.png",
# #     "table": []
# #   }}
# # ]
# # Now generate the slide data.
# """


ENRICH_PRESENTATION_PROMPT = """
You are generating a presentation from a document.
Return the slides as a JSON array.
Each slide should include:

{{
  "title": "Slide title",
  "text": "• bullet 1\\n• bullet 2",
  "image": "path/to/image.png",
  "table": [["Header 1", "Header 2"], ["Row 1 Col 1", "Row 1 Col 2"]]
}}

Document text:
{text_content}

Available images with context:
{image_paths}

Available table data:
{table_data}

# Guidelines:
# - Each slide should have 4-6 concise bullet points, each containing a single key idea or fact.
# - **The number of slides should be determined by the document's topics and subtopics, not limited by the number of images.**
# - For each slide, focus on a distinct topic or subtopic from the document. 
# - **Split long sections into multiple slides if they contain many bullet points, multiple concepts, or detailed breakdowns.**
# - Split content into as many distinct slides as needed based on content structure. Do not limit to 8–10 slides.
# - When splitting content into multiple slides, avoid generic titles like "Part 1", "Part 2", or "continued" unless no better option is possible.
# - Instead, try to generate unique, descriptive slide titles that clearly reflect the key idea, subtopic, or focus of the slide.
# - **Do not crowd a slide with too much text. Prefer multiple shorter slides over one long one.**
# - **Also don't oversimplify slides with just 2-3 word bullets where explanation could be required.**
# - Generate multiple slides for major sections if needed, breaking down content into digestible chunks.
# - If a table is included, do not break it across multiple slides unless multiple unrelated topics are covered.
# - If all rows share the same structure and relate to one concept (e.g., AI use cases by domain), include the **entire table on a single slide**.
# - Only split a table across slides if it is very large **and** spans clearly different topics.
# - Output must include the 'table' field for every slide. Leave it empty ("table": []) only if no structured data is relevant.
# - If a table is shown on a slide, do not repeat its contents in the bullet points. Avoid duplicating the same information across the "text" and "table" fields.
# - Do not break the same table into multiple slides. Consolidate all related rows in a single table and place it on one slide.
# - Only create a new slide for a table if the table is clearly associated with a distinct section or topic.
# - If you include a table on a slide, do not repeat the same information in bullet points. Choose one format per piece of content.
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
"""

EXTRACT_TOPICS_MARKERS_TEMPLATE = """
You are an expert document analyst.

Analyze the following document and extract **all meaningful topics or subtopics**, following these rules:

1. Extract any section, paragraph, or idea that could be shown as its own slide.
2. Topics should reflect distinct concepts, domains, or shifts in the document focus.
3. Aim to extract 5–7 key topics minimum unless the document is extremely short.
4. Avoid repeating topics unless clearly expanded upon later.


Provide the output in the following format:
**<key topic 1>**
first ten words of the section or theme that the key topic 1 represents
**<key topic 2>**
first ten words of the section or theme that the key topic 2 represents

Do NOT add any explanations.

Now extract from the document below:
Document:
'''
{{content}}
'''
"""



# EXTRACT_TOPICS_MARKERS_TEMPLATE = """
# Analyze the document and extract well-defined topics or subtopics. Aim to extract 5–7 key topics minimum unless the document is extremely short.
# Use the following logic:

# 1. Each topic should correspond to a unique theme or section, not just broad categories.
# 2. Use clues like section headings, transitions, or examples that shift focus.
# 3. Include subtopics where applicable (e.g., separate Education vs. Healthcare if both use AI).
# 4. For each topic, include:
#    • Topic Title
#    • First 10–12 words from the section start as a sample

# Output format:
# Topic: <title>
# Sample: <first few words>

# Document:
# {{content}}
# """



# EXTRACT_TOPICS_MARKERS_TEMPLATE = """
# Analyze the given document and extract key topics, following these guidelines:

# 1. Key Topic Identification:
#    - Topics should represent major sections or themes in the document.
#    - Each key topic should be substantial enough for at least one slide with 3-5 bullet points, potentially spanning multiple slides.
#    - Topics should be broad enough to encompass multiple related points but specific enough to avoid overlap.
#    - Identify topics in the order they appear in the document.
#    - Consider a new topic when there's a clear shift in the main subject, signaled by transitional phrases, new headings, or a distinct change in content focus.
#    - If a topic recurs, don't create a new entry unless it's substantially expanded upon.

# 2. Key Topic Documentation:
#    - For each key topic, create a detailed name that sums up the idea of the section or theme it represents. 
#    - Next, provide the first ten words of the section that the key topic represents.

# 3. Provide the output in the following format:
# **<key topic 1>**
# first ten words of the section or theme that the key topic 1 represents
# **<key topic 2>**
# first ten words of the section or theme that the key topic 2 represents

# Document to analyze:
# '''
# {{content}}
# '''

# """

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


SLIDE_REFINEMENT_PROMPT = """
You are refining a structured presentation generated from a document.

Each slide contains:
- A "title"
- A "text" field (bullet points)
- An optional "image"
- An optional "table" (list of rows and columns)

Your task is to enhance the overall quality of the slides following these rules:

1. If a slide contains both a table and bullet points, avoid repeating the same information in both.
   - Use bullets to provide context, explanation, or additional perspectives.

2. Across the entire deck, avoid repeating the same information multiple times across different slides.
   - If multiple slides cover overlapping table rows or bullet points, adjust later slides to focus on different aspects.
   - Prefer expansion, contrast, or new details instead of repeating facts already covered earlier.

3. Improve bullet point quality:
   - Make each bullet precise, informative, and fact-driven.
   - Avoid vague or generic buzzwords.

4. Improve presentation flow:
   - Ensure logical progression of ideas across slides.
   - Reword titles slightly if needed for clarity and smooth transitions.

5. Maintain formatting:
   - Keep the JSON structure exactly the same.
   - Modify only the "title" and "text" fields.
   - Do not change or remove "table" or "image" fields.

Return a valid JSON array of the refined slides.

Here is the original slide JSON:
{original_slide_json}
"""