#!/usr/bin/env python3

import os
import requests
import openai
import glob
import webbrowser
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# --- MANDATORY SETTINGS ---

# Set your API key, and paths for PowerPoint template and output directory.
openai.api_key = "YOUR-API-KEY"
template_path = "/path/of/template.pptx" # default template path must be set
output_directory = "/path/to/output/directory/"
template_folder = "/home/gat/ppt/templates/"

# Set slides output language
language = "English"

# If you are not using English, set to False to translate prompt for DALL-E
english = True

# Picture prompt. Change last part to your favourite style, MUST be in English.
IMAGE_PROMPT = "a portrait photo of {topic}, detailed, cgi, octane, unreal"

# Define temperature value. Determines GPT randomness.
keypoint_temperature_value = 0.5
content_temperature_value = 0.7

# ---- OTHER SETTINGS ----
# Edit only if you know what to do with gpt prompts!

KEYPOINT_PROMPT = "Write 8 short titles of maximum 6 words on key points to be covered in a lesson about topic: {topic}. It is important that terms {topic} always appear. From now on we will interact ONLY in good {language} language."
CONTENT_PROMPT = "Summarize in 6 points and using at least 10 words the most important aspects of following topic: {topic}.\nShow me ONLY the list and no other text. From now on we will interact ONLY in good {language} language."

# Define token length. Determines request and response lenght.
keypoint_max_tokens=2000
content_max_tokens=2000

# do not edit this lines
template_path = os.path.join(template_folder, "template.pptx") 
image_size = "1024x1024"

# --- CONFIGURATION END ---

def get_template_filenames():
    template_files = glob.glob(os.path.join(template_folder, "*.pptx"))
    return [os.path.basename(f) for f in template_files]
    
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
    topic_name = topic_entry.get()
    prompt = KEYPOINT_PROMPT.format(topic=topic_name, language=language)
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
        prompt = CONTENT_PROMPT.format(topic=topic, language=language)
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
    gpt_outputs = []  # Add this line to create an empty list for storing GPT outputs.

    for topic, text in lists:
        gpt_outputs.append((topic, text))  # Add this line to store the GPT outputs.

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
    
    # Send the gpt_outputs list to GPT-3 to generate an essay.
    essay_prompt = "Please write an essay based on the following points:\n"
    for topic, text in gpt_outputs:
        essay_prompt += f"\n{topic}\n{text}\n"    
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=essay_prompt,
        max_tokens=2000,
        n=1,
        stop=None,
        temperature=0.7,
    )  

    # Generate and download an image using DALL-E with topic in English.
    for i in range(num_images):
        image_url = "https://api.openai.com/v1/images/generations"
        prompt = IMAGE_PROMPT.format(topic=topic_name_en)
        payload = {
            "model": "image-alpha-001",
            "prompt": prompt,
            "num_images": 1,
    	    "size": image_size,
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

def generate_images():
        
    # Get topic name and number of images from input fields.
    topic_name = topic_entry.get()
    topic_name_en = topic_entry.get()
    num_images = int(num_images_entry.get())
    if not english:
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

    # Generate and download only images using DALL-E with topic in English.
    for i in range(num_images):
        image_url = "https://api.openai.com/v1/images/generations"
        prompt = IMAGE_PROMPT.format(topic=topic_name_en)
        payload = {
            "model": "image-alpha-001",
            "prompt": prompt,
            "num_images": 1,
 	       "size": image_size,
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
        pptx_filename = os.path.join(output_directory, f"{topic_entry.get()}.pptx")
        image_filename = os.path.splitext(pptx_filename)[0] + f"_image_{num_images + i + 1}.png"
        with open(image_filename, "wb") as f:
            f.write(image_content)
        update_status_bar(f"Generated image {num_images + i + 1} of {num_images * 2}!")
                 
# Create main window.
window = tk.Tk()
window.title("G-PoinT")

# Set window size.
window_width = 350
window_height = 400
window.geometry(f"{window_width}x{window_height}")

# Center window.
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x = int((screen_width/2) - (window_width/2))
y = int((screen_height/2) - (window_height/2))
window.geometry(f"+{x}+{y}")

# Create a label for topic entry ("Presentation topic").
topic_label = tk.Label(window, text="Presentation topic:")
topic_label.pack()

# Create text entry for topic.
topic_entry = tk.Entry(window)
topic_entry.pack()

# Create label for number of pictures entry ("Pictures to be generated").
num_images_label = tk.Label(window, text="How many pictures:")
num_images_label.pack()

# Create text entry for number of images.
num_images_entry = tk.Entry(window)
num_images_entry.insert(0, "1")
num_images_entry.pack()

# Create a label for image size selection.
image_size_label = tk.Label(window, text="Picture size:")
image_size_label.pack()

# Define a variable to hold the selected image size.
image_size_var = tk.StringVar(window)
image_size_var.set("1024x1024")  # Set default value.

# Define the available image sizes as a list of strings.
image_sizes = ["256x256", "512x512", "1024x1024"]

# Create the dropdown menu for selecting the image size.
image_size_menu = tk.OptionMenu(window, image_size_var, *image_sizes)
image_size_menu.pack()

# Create a label for template selection.
template_label = tk.Label(window, text="Select template:")
template_label.pack()

# Define a variable to hold the selected template.
template_var = tk.StringVar(window)
template_var.set("template.pptx")  # Set default value.

# Get the list of available templates.
templates = get_template_filenames()

# Create the dropdown menu for selecting the template.
template_menu = tk.OptionMenu(window, template_var, *templates)
template_menu.pack()

# Create a function to update the template_path variable based on the selected value.
def update_template_path():
    global template_path
    selected_template = template_var.get()
    template_path = os.path.join(template_folder, selected_template)

# Bind the update_template_path function to the dropdown menu.
template_var.trace("w", lambda *args: update_template_path())

# Create a function to update the image_size variable based on the selected value.
def update_image_size():
    global image_size
    image_size = image_size_var.get()

# Bind the update_image_size function to the dropdown menu.
image_size_var.trace("w", lambda *args: update_image_size())
separator = ttk.Separator(window, orient='horizontal')
separator.pack(fill='x', padx=10, pady=10)

# Create button to generate PowerPoint and image.
generate_button = tk.Button(window, text="Generate presentation and pictures", command=generate_powerpoint_and_image)
generate_button.pack(padx=10, pady=10)
generate_more_button = tk.Button(window, text="Generate more (or only) pictures!", command=generate_images)
generate_more_button.pack(padx=2, pady=2)

# Create a label to display the template path and output directory.
path_label = tk.Label(window, text=f"Templates: {template_folder}\nOutput: {output_directory}")
path_label.pack(padx=10, pady=10)

# Create status bar label.
status_bar = tk.Label(window, text="Push and wait for notice...", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)
def show_help_message():
    message = f"Presentation topic: fill with desired topic, word or phrase (e.g. Dolphin or Mona Lisa Story).\n\nHow many pictures: number of pictures to be generated by DALL-E. Default is 1.\n\nPicture size: select desired picture size from dropdown menu, choose between: 256x256, 512x512, 1024x1024.\n\nGenerate presentation and pictures: click to generate a full PowerPoint presentation with text, and download pics. Generated files will be saved in the output directory specified in the script.\n\nGenerate more pictures: click to get additional (n) pics without PowerPoint presentation..\n\nNote: you must have a valid OpenAI API key, and set correct template and output paths in the script before running."
    tk.messagebox.showinfo("Help", message)
    return

def show_about_message():
    about_window = tk.Toplevel(window)
    about_window.title("About")
    about_window.geometry("400x200")
    
    message = "G-PoinT v1.2.0 by gat\n\nGenerates complete PowerPoint files and images, using OpenAI's GPT-3 API and DALL-E image generation model.\n\nhttps://github.com/davidegat/G-PoinT\n\nLicensed under the GNU General Public License v3.0"
    
    text_widget = tk.Text(about_window, wrap=tk.WORD, padx=10, pady=10)
    text_widget.insert(tk.END, message)
    text_widget.pack(expand=True, fill=tk.BOTH)

    text_widget.tag_configure("hyperlink", foreground="blue", underline=1)
    text_widget.tag_bind("hyperlink", "<Button-1>", lambda e: webbrowser.open("https://github.com/davidegat/G-PoinT"))

    start_index = text_widget.search("https://github.com/davidegat/G-PoinT", tk.END)
    end_index = text_widget.index(f"{start_index}+{len('https://github.com/davidegat/G-PoinT')}c")

    text_widget.tag_add("hyperlink", start_index, end_index)
    text_widget.config(state=tk.DISABLED)
    # Center about_window.
    about_window_width = 400
    about_window_height = 200
    about_window_x = window.winfo_x() + int((window_width / 2) - (about_window_width / 2))
    about_window_y = window.winfo_y() + int((window_height / 2) - (about_window_height / 2))
    about_window.geometry(f"+{about_window_x}+{about_window_y}")

# Define function to handle exit button click.
def exit_application():
    window.destroy()
    
menu_bar = tk.Menu(window)
menu_menu = tk.Menu(menu_bar, tearoff=0)
menu_menu.add_command(label="Exit", command=exit_application)
menu_bar.add_cascade(label="Menu", menu=menu_menu)
menu_bar.add_command(label="Help", command=show_help_message)
menu_bar.add_command(label="About", command=show_about_message)

window.config(menu=menu_bar)

# Start main loop.
window.mainloop()
