# Doc2Deck

SmartSlideGen is an AI-powered pipeline that automatically converts Microsoft Word documents into professional PowerPoint presentations. The system extracts text and embedded images in the order they appear, processes the text via a generative AI API to create slide content, and dynamically integrates images with the corresponding text to produce a cohesive, visually appealing presentation.

## Features

- **Text Extraction:**  
  Extracts text sequentially from a Word document using the `python-docx` library.

- **Image Extraction:**  
  Identifies and extracts embedded images, saving them locally with unique filenames.

- **AI-Powered Slide Generation:**  
  Generates structured slide content (titles and bullet points) by sending custom prompt templates to an external AI API.

- **Dynamic Slide Creation:**  
  Uses the `pptx` library to load a PowerPoint template, clear existing slides, and create new slides with both text and images.

- **Content Filtering:**  
  Implements skip phrases and duplicate detection (via Python’s `SequenceMatcher`) to avoid redundant slide creation.

- **Error Handling:**  
  Incorporates robust error management during API calls and image processing to ensure smooth execution.

