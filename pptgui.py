#!/usr/bin/env python3

import os
import tkinter as tk
from pptx import Presentation
from pptx.util import Inches, Pt
import requests
from PIL import Image
from io import BytesIO
from pptx.dml.color import RGBColor

# Import the OpenAI API and set the API key
import openai
openai.api_key = "YOUR API KEY"

# Set the path for the template PowerPoint file
template_path = "/home/gat/ppt/template.pptx"

# Define a function to generate the PowerPoint presentation and image
def generate_powerpoint_and_image():
    # Get the topic name from the input field
    topic_name = topic_entry.get()

    # Translate the topic name to English using GPT for DALL-E prompt
    prompt = f"Translate {topic_name} to English."
    response = openai.Completion.create(
        engine="text-davinci-002",
        prompt=prompt,
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.7,
    )
    topic_name_en = response.choices[0].text.strip()
    
   # Generate 8 key points to use as slides
    prompt = f"Scrivi 8 brevi titoli di massimo 6 parole di punti fondamentali da trattare in una lezione su {topic_name}. È importante che compaiano sempre i termini: {topic_name}."
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        max_tokens=200,
        n=1,
        stop=None,
        temperature=0.5,
    )
    topic_prompts = response.choices[0].text.strip()

    # Save to a text file
    with open('topics.txt', 'w') as f:
        f.write(topic_prompts)

    # Generate the slides contents using the OpenAI GPT-3 API - Translate prompt in your own language
    
    with open('topics.txt', 'r') as f:
        topics = f.read().splitlines()

    lists = []
    for topic in topics:
        prompt = f"Riassumi in 6 punti e usando MINIMO QUINDICI PAROLE gli aspetti più importanti del seguente argomento: {topic}\nNon aggiungere avvisi e scrivi solo l'elenco."
        response = openai.Completion.create(
            engine="text-davinci-003",
            prompt=prompt,
            max_tokens=2000,
            n=1,
            stop=None,
            temperature=0.7,
        )
        text = response.choices[0].text.strip()
        lists.append((topic, text))

    # Create a new PowerPoint presentation from the template
    prs = Presentation(template_path)
    
    for topic, text in lists:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = topic
        title_shape = slide.shapes.title
        title_shape.text = topic
        title_shape.text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].font.size = Pt(42)

        left = Inches(1)
        top = Inches(2.5)
        width = Inches(9)
        height = Inches(4)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        for item in text.split("\n"):
            p = tf.add_paragraph()
            p.text = item.strip()
            p.level = 0
            p.font.bold = False
            p.font.size = Inches(0.35)
            p.font.color.rgb = RGBColor(0, 0, 0)

    # Save the PowerPoint presentation to desktop - you may want to customize
    pptx_filename = f"/home/gat/Scrivania/{topic_entry.get()}.pptx"
    prs.save(pptx_filename)
    # Generate an image using the DALL-E API

    image_url = "https://api.openai.com/v1/images/generations"
    prompt = f"a portrait photo of {topic_name_en}, detailed, cgi, octane, unreal"
    payload = {
        "model": "image-alpha-001",
        "prompt": prompt,
        "num_images": 1,
        "size": "1024x1024",
        "response_format": "url"
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {openai.api_key}"
    }
    response = requests.post(image_url, headers=headers, json=payload)
    response.raise_for_status()
    image_url = response.json()["data"][0]["url"]
    image_content = requests.get(image_url).content

    # Save the image to a file in desktop
    image_filename = os.path.splitext(pptx_filename)[0] + ".png"
    with open(image_filename, "wb") as f:
        f.write(image_content)
        
# Create the main window
window = tk.Tk()
window.title("G-PoinT")

# Create a label for the topic entry
topic_label = tk.Label(window, text="Argomento della presentazione:")
topic_label.pack()

# Create a text entry for the topic
topic_entry = tk.Entry(window)
topic_entry.pack()

# Create a button to generate the PowerPoint and image
generate_button = tk.Button(window, text="Genera file PowerPoint e immagine", command=generate_powerpoint_and_image)
generate_button.pack()

# Start the main event loop
window.mainloop()

