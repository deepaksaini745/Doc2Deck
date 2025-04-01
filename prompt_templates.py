# prompt_templates.py

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