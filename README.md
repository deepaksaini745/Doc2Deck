# 📚 Doc2Deck – AI-Powered Document-to-Presentation Generator

**Doc2Deck** is an end-to-end Python pipeline that converts unstructured `.docx` documents into structured, visually engaging PowerPoint presentations. It combines document parsing, GPT-based topic generation, image captioning, and table reconstruction to automate professional slide creation with minimal manual effort.

---

## 🚀 Features

- 📄 **Advanced Document Parsing** using `LlamaParse` for extracting text, tables, and images.
- 🧠 **Topic-Aware Chunking** of long documents with GPT-based semantic grouping.
- 🖼️ **Image Filtering & Captioning** to skip blank/solid images and generate smart captions via GPT.
- 📊 **Table Handling** with fallback to Markdown extraction if structured data is missing.
- 🎯 **Image-to-Slide Mapping** using fuzzy logic and GPT-based caption matching.
- 🧹 **Content Refinement** to clean up GPT output and avoid repetition or filler slides.
- 🎨 **PowerPoint Generation** with custom formatting, layout adjustments, and templating support.

---

## 🧩 Tech Stack

- **Python 3.8+**
- [python-pptx](https://github.com/scanny/python-pptx)
- [LlamaParse (LlamaIndex Cloud)](https://llamahub.ai/)
- [OpenAI GPT-4 / GPT-3.5 Turbo](https://openai.com/)
- [fuzzywuzzy](https://github.com/seatgeek/fuzzywuzzy)
- [Pillow](https://pillow.readthedocs.io/)
- [dotenv](https://pypi.org/project/python-dotenv/)

