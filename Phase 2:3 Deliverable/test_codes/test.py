import os
from llama_cloud_services import LlamaParse


base_dir = os.path.dirname(os.path.abspath(__file__))
print("base_dir = ", base_dir)

doc_path = os.path.join(base_dir, "input", "doc.docx")

LLAMA_CLOUD_API_KEY = os.getenv("LLAMA_CLOUD_API_KEY")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

parser = LlamaParse(api_key=LLAMA_CLOUD_API_KEY, language="en", verbose=True)
result = parser.parse(doc_path)

images = result.get_image_documents(
    include_screenshot_images=True,
    include_object_images=True,
    image_download_dir="./pptWithBreaking/images"
)

images_dir = os.path.join(base_dir, "images")
print("images_dir = ", images_dir)


# image_paths = [img.file_path for img in images if hasattr(img, "file_path") and os.path.exists(img.file_path)]

image_paths = [
    os.path.join(images_dir, f)
    for f in os.listdir(images_dir)
    if f.lower().endswith(('.png', '.jpg', '.jpeg'))
]

print(f"Collected image paths from directory: {image_paths}")

# print("image_paths = ", image_paths)


