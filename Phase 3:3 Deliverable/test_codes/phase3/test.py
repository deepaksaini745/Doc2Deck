import os
import openai
openai.api_key = os.getenv("OPENAI_API_KEY")

import base64
import mimetypes

def get_image_caption(image_path):
    mime_type, _ = mimetypes.guess_type(image_path)
    with open(image_path, "rb") as image_file:
        image_data = base64.b64encode(image_file.read()).decode("utf-8")
    
    # Create a data URL
    data_url = f"data:{mime_type};base64,{image_data}"

    response = openai.ChatCompletion.create(
        model="gpt-4-turbo",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "Describe this image in 2-3 sentences focusing on objects and key ideas."},
                    {"type": "image_url", "image_url": {"url": data_url}}
                ]
            }
        ],
        max_tokens=300
    )
    caption = response['choices'][0]['message']['content']
    return caption

print(get_image_caption("/Users/deepaksaini/Desktop/ETB_Project/pptWithBreaking/images/img_p3_3_FIG__1_Two_authors_independently_analysed_the_1877.png"))
