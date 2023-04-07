# G-PoinT
![gpoint](https://user-images.githubusercontent.com/51516281/230445369-7ec377ca-d642-4dce-aebd-b1b3e2b7c6d6.png)
<br>
This software uses Python3 and GPT API to create a <b>complete</b> PowerPoint file, including slides and text, <b>from a single topic input</b>. DALL-E API is then used to generate and download an <b>appropriate image</b> to use within the presentation. Only tested on Linux 5.15.0-69 / Ubuntu 20.04.1 / Python 3.8.10. You will need to translate few GUI elements to your language (currently Italian) and customize paths. Code is adequately commented, with English prompts and instructions provided where necessary. 

<h3>Install and run</h3>
<b>Download:</b> <code>git clone https://github.com/davidegat/G-PoinT.git</code><br>
<b>Python3 dependencies</b>: <code>pip install python-pptx requests Pillow openai</code><br>
<b>pptgui.py configuration</b>:<br>
<code>openai.api_key = "your_api_key"
template_path = "/path/to/your/template.pptx"
output_directory = "/path/to/your/output/directory"</code><br><br>
<b>Running and using G-PoinT</b>:<br>
<code>cd G-PoinT</code><br>
<code>python3 ./pptgui.py</code><br>
You can use the included G-PoinT.desktop file and access GUI via desktop, remember to edit paths accordingly and make it executable: <code>chmod +x G-PoinT.desktop</code><br><br>

Insert a topic and click on generate button (for example, "Brain Tumor"), wait for a reasonable amount of time to receive PPTX and PNG outputs directly in custom folder. <b>Please note that it may take up to one minute to generate both PowerPoint and image!</b><br><br>
This repository contains <b>examples of output</b> and an <b>example template</b> you can customize or replace with your own.
<h3>Results from GPT and DALL-E</h3>
To obtain different results (more slides, more text, specific contexts), modify the prompt sent to GPT into the script directly. Try different prompts, temperatures, and tokens for fine-tuned results. If you need more realistic, artistic, or other styles for image generation, modify DALL-E prompt accordingly, read comments in code for details. Also refer to <a href="https://python-pptx.readthedocs.io/en/latest/">pptx library documentation</a> to customize font, colors, text size and other presentation elements, or if you want to include image directly into presentation.<br><br>
<b>The following video shows process in real time.</b> Please wait until files are complete, or skip from 0:15 to 0:55.<br><br>

[https://github.com/davidegat/G-PoinT/raw/main/G-Point.mp4](https://user-images.githubusercontent.com/51516281/230441862-0ee64ad7-b564-49d1-beba-8e0fb085b885.mp4)

<h3>What it does?</h3>

<li>Shows Tkinter GUI and asks user for a topic
<li>Sends prompt and topic to GPT
<li>Gets back 8 key points, each one will be a slide
<li>Sends one prompt for each key point to generate slide content text
<li>Creates a PPTX file from template with key points as slide titles, and fills them with generated text
<li>Translates topic from whatever language to english (also via GPT prompt)
<li>Asks DALL-E for an image using topic in english
<li>Saves both PPTX and PNG files.
<li>First slide will be empty for user customization.


<h3>Dependencies and requirements</h3>

<li>os - a built-in Python library for interacting with the operating system.
<li>tkinter - a built-in Python library for creating graphical user interfaces (GUIs).
<li>pptx - a Python library for creating and updating PowerPoint (.pptx) files.
<li>requests - a popular Python library for making HTTP requests.
<li>PIL (Python Imaging Library) - an open-source Python library for adding image processing capabilities to your Python interpreter.
<li>openai - the official Python library for the OpenAI API, used to interact with OpenAI services like GPT-3 and DALL-E.

Please note that the script assumes you have a Linux box and compatible version of Python 3 installed on your system (tested on Python 3.8.10). Additionally, script relies on having access to the OpenAI API key, which you'll need to $ign up for.
<h3>TO-DO</h3>
<li>Generate more images, allowing to choose the best ones.
<li>Insert image(s) in PPTX directly.
<li>Move prompts to external files for easier customization.
<li>Other platform (e.g. Windows) testing.

<h3>Not professional. Not perfect.</h3>
G-PoinT is not a substitute for a real person in terms of content, but if you know what to do with prompts you could get good starting points to speed up your workflow, and get inspired. Try to play around with templates and pptx library for more impressive, yet not perfect, graphic results.
