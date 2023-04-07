# G-PoinT
![gpoint](https://user-images.githubusercontent.com/51516281/230445369-7ec377ca-d642-4dce-aebd-b1b3e2b7c6d6.png)
<br>
This software uses Python3 and GPT to create a <b>complete</b> PowerPoint file, which includes slides and text, <b>from a single topic input</b>. DALL-E is then utilized to generate and download an appropriate image to use within the presentation. Software is released AS-IS: you will only need to translate the GUI to your language (currently Italian) and customize paths. Code is adequately commented, with English prompts and instructions provided where necessary.

<h3>Install and run</h3>
<b>Download:</b> <code>git clone https://github.com/davidegat/G-PoinT.git</code><br>
<b>Dependencies</b>: <code>pip install python-pptx requests Pillow openai</code><br><br>
<b>Running G-PoinT</b>:<br>
<li><code>cd G-PoinT</code><br>
<li><code>python3 ./pptgui.py</code><br>
Or also:
<li><code>cd G-PoinT</code><br>
<li><code>chmode +x ./pptgui.py</code> (only once)
<li><code>./pptgui.py</code><br><br>

Insert a topic and click on the generate button (for example, "Brain Tumor"), and wait for a reasonable amount of time to receive PPTX and PNG outputs directly in custom folder. <b>Please note that it may take up to one minute to generate both PowerPoint and image!</b><br><br>
This repository contains <b>examples of output</b>, an <b>example template</b> you can customize, and a <b>.desktop file</b> for desktop entry (which must be edited).<br><br>
To obtain different results (more slides, more text, specific contexts), modify the prompt sent to GPT. Try different prompts, temperatures, and tokens for fine-tuned results. If you want more realistic, artistic, or other styles for images, modify DALL-E prompt accordingly. Please refer to the comments for details.<br><br>
The following video shows the process in real time. Please wait until the files are complete, or skip to 0:55 from 0:15.

[https://github.com/davidegat/G-PoinT/raw/main/G-Point.mp4](https://user-images.githubusercontent.com/51516281/230441862-0ee64ad7-b564-49d1-beba-8e0fb085b885.mp4)

<h3>What it does?</h3>

<li>Shows Tkinter GUI and asks for a topic
<li>Sends prompt + topic to GPT
<li>Gets back 8 key points, each one will be a slide
<li>Sends another prompt to generate text for each slide
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

Please note that the script assumes you have a compatible version of Python 3 (preferably Python 3.6 or later) installed on your system. Additionally, the script relies on having access to the OpenAI API key, which you'll need to sign up for.
<h3>Not professional. Not perfect.</h3>
