#!/usr/bin/env python3

import os
import requests
import openai
import tkinter as tk
from tkinter import messagebox
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# --- CONFIGURATION ---

# English Language? set to True. Remember to use English prompt instead!
english = True

# Set your API key, and paths for PowerPoint template and output directory.
openai.api_key = "your_api_key"
template_path = "/path/to/your/template.pptx"
output_directory = "/path/to/your/output/directory"

# Define prompts text (keypoint_prompt = slides number and titles, content_prompt = slides content).
KEYPOINT_PROMPT = "Scrivi 8 brevi titoli di massimo 6 parole di punti fondamentali da trattare in una lezione su {topic}. È importante che compaiano sempre i termini: {topic}."
# In english: Write 8 short titles of maximum 6 words on key points to be covered in a lesson on {topic}. It is important that terms {topic} always appear.

CONTENT_PROMPT = "Riassumi in 6 punti e usando MINIMO QUINDICI PAROLE gli aspetti più importanti del seguente argomento: {topic}\nNon aggiungere avvisi e scrivi solo l'elenco."
# In English: Summarize in 6 points and using a minimum of fifteen words the most important aspects of following topic: {topic}. Do not add warnings and write only list.

IMAGE_PROMPT = "a portrait photo of {topic}, detailed, cgi, octane, unreal"

# Define temperature value. Determines GPT randomness.
keypoint_temperature_value = 0.5
content_temperature_value = 0.7
        
# Define token length. Determines request and response lenght.
keypoint_max_tokens=2000
content_max_tokens=2000

# --- CONFIGURATION END ---

# Define update_status_bar function.
def update_status_bar(text):
    status_bar.config(text=text)
    
# Define a function to generate PowerPoint presentation and image.
def generate_powerpoint_and_image():
    # Get topic name and number of images from input fields.
    topic_name = topic_entry.get()
    topic_name_en = topic_entry.get()
    num_images = int(num_images_entry.get())
    if not english:
        # execute code if non english language
        # Translate topic to English with GPT, to build DALL-E prompt later.
        prompt = f"Translate {topic_name} to English."
        response = openai.Completion.create(
            engine="text-davinci-002",
            prompt=prompt,
            max_tokens=200,
            n=1,
            stop=None,
            temperature=0.5,
        )
        topic_name_en = response.choices[0].text.strip()

    # Generate key points to use as slides titles and to generate text later.
    topic_prompts = ""
    prompt = KEYPOINT_PROMPT.format(topic=topic_name)
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        max_tokens=keypoint_max_tokens,
        n=1,
        stop=None,
        temperature=keypoint_temperature_value,
    )

    topic_prompts = response.choices[0].text.strip()
    lists = []
    for topic in topic_prompts.split("\n"):
    
        # Generate text content for each key point.
        text = ""
        prompt = CONTENT_PROMPT.format(topic=topic)
        response = openai.Completion.create(
            engine="text-davinci-003",
            prompt=prompt,
            max_tokens=content_max_tokens,
            n=1,
            stop=None,
            temperature=content_temperature_value,
        )
        text = response.choices[0].text.strip()
        lists.append((topic, text))
        
    # Create a new PowerPoint presentation from template
    # and fill it with generated key points (slides) and text.
    # See 'pptx' library documentation for further graphic customization.
    prs = Presentation(template_path)

    for topic, text in lists:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = topic
        title_shape = slide.shapes.title
        title_shape.text = topic

        # Format title font.
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
            
            # Format content font.
            p.font.bold = False
            p.font.size = Inches(0.35)
            p.font.color.rgb = RGBColor(0, 0, 0)
            
    # Save presentation.
    pptx_filename = os.path.join(output_directory, f"{topic_entry.get()}.pptx")
    prs.save(pptx_filename)             
    # Generate and download an image using DALL-E with topic in English.
    for i in range(num_images):
        image_url = "https://api.openai.com/v1/images/generations"
        prompt = IMAGE_PROMPT.format(topic=topic_name_en)
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

        # Save image.
        image_filename = os.path.splitext(pptx_filename)[0] + f"_image_{i+1}.png"
        with open(image_filename, "wb") as f:
            f.write(image_content)
        update_status_bar("Files generated!") 
                 
# Create main window.
window = tk.Tk()
window.title("G-PoinT")

# Set window size.
window_width = 400
window_height = 200
window.geometry(f"{window_width}x{window_height}")

# Center window.
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x = int((screen_width/2) - (window_width/2))
y = int((screen_height/2) - (window_height/2))
window.geometry(f"+{x}+{y}")

# Create a label for topic entry ("Presentation Topic").
topic_label = tk.Label(window, text="Argomento della presentazione:")
topic_label.pack()

# Create text entry for topic.
topic_entry = tk.Entry(window)
topic_entry.pack()

# Create label for number of pictures entry ("Pictures to be generated").
num_images_label = tk.Label(window, text="Immagini da generare:")
num_images_label.pack()

# Create text entry for number of images.
num_images_entry = tk.Entry(window)
num_images_entry.insert(0, "1")
num_images_entry.pack()

# Create button to generate PowerPoint and image ("Generate file(s)!").
generate_button = tk.Button(window, text="Genera i file!", command=generate_powerpoint_and_image)
generate_button.pack()

# Create status bar label.
status_bar = tk.Label(window, text="Push and wait for notice...", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# Create menu and menu bar items
def show_about_message():
    message = "G-PoinT v1.2.0 by gat"
    tk.messagebox.showinfo("About", message)

# Define function to handle the exit button click event.
def exit_application():
    window.destroy()
    
menu_bar = tk.Menu(window)
menu_menu = tk.Menu(menu_bar, tearoff=0)
menu_menu.add_command(label="Exit", command=exit_application)
menu_bar.add_cascade(label="Menu", menu=menu_menu)
menu_bar.add_command(label="About", command=show_about_message)

window.config(menu=menu_bar)

# Start main event loop.
window.mainloop()
