# Doc2Deck

**Doc2Deck** is an AI-powered document-to-presentation pipeline that automates the conversion of unstructured Microsoft Word documents into structured, professionally formatted PowerPoint slides. Designed for analysts, consultants, educators, and business professionals, it streamlines the traditionally manual and repetitive slide creation process using modern language models and document parsers.

---

## Key Features

* **Structured Extraction**: Parses `.docx` files to extract paragraphs, images, and tables using **LlamaParse**.
* **LLM-Driven Slide Generation**: Uses **Claude** and **GPT APIs** to generate titles, bullet points, and layout suggestions.
* **Intelligent Image Mapping**: Associates images with relevant slide content using captioning and semantic similarity scoring.
* **Table Rendering Support**: Extracts and formats both structured and Markdown-style tables into clean PowerPoint slides.
* **Design Consistency**: Ensures visual uniformity with predefined templates (fonts, spacing, layout hierarchy).
* **Pipeline Orchestration**: Modular code for document parsing, prompt generation, layout logic, and output rendering.
* **Robust Error Handling**: Fallback logic for API failures, schema mismatches, and layout inconsistencies.

---

## How It Works

### 1. Document Parsing

* Extracts text, images, and tables using **LlamaParse**.
* Renames visual assets using contextual metadata.

### 2. Topic-Aware Chunking

* Uses **GPT-4** to segment document text into semantically meaningful sections.

### 3. Content Enrichment

* Generates slide-level JSON (titles, bullets, layout hints) via **Claude** or **GPT-4**.
* Applies refinement logic for formatting and consistency.

### 4. Image Captioning and Mapping

* **GPT-4 Turbo** generates captions for images.
* Semantic matching ensures images are placed with the most relevant slides.

### 5. Table Insertion

* Dynamically renders tables with bolded headers and adaptive layouts.
* Falls back to Markdown parsing if needed.

### 6. PowerPoint Rendering

* Uses **python-pptx** to generate the final slide deck from enriched JSON.
* Includes custom "Thank You" slide with branded formatting.

---

## Technology Stack

* **Python** 3.10+
* [**LlamaParse**](https://llamaindex.ai/) (LlamaIndex API) – for document parsing
* **Claude** (Anthropic API) – for slide generation and JSON structuring
* **GPT-4 / GPT-4 Turbo** (OpenAI API) – for topic chunking, image captioning, and layout refinement
* **python-docx** – for legacy `.docx` support
* **python-pptx** – for PowerPoint creation
* **FuzzyWuzzy** / `difflib` – for image-slide similarity scoring and deduplication

---

## Known Challenges

* Inconsistent JSON structures from LLMs across large documents
* Caption-slide mismatches due to low relevance or ambiguity
* Table rendering issues with overly wide or complex structures
* API rate limits and token constraints on long documents

---

## Future Roadmap

* Web-based interface for real-time slide previews
* Embedding-based visual mapping using vector similarity
* Auto-diagram generation using **Mermaid** or **Graphviz**
* Enterprise-grade CI/CD pipeline for scalable deployment

---

## Contributors

* **Deepak Saini** – [saini50@purdue.edu](mailto:saini50@purdue.edu)
* **Mokshda Sharma** – [sharm879@purdue.edu](mailto:sharm879@purdue.edu)
