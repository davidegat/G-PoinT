# G-PoinT
![gpoint](https://user-images.githubusercontent.com/51516281/230658427-1af87422-e6f1-4b42-bfdd-ab1f411a1add.png)
<br><br>
This software uses Python3 and GPT via <a href="https://platform.openai.com/docs/api-reference/introduction">OpenAI API</a> to create a <b>complete</b> PowerPoint file from template, including slides and text, <b>from a single topic input</b>. <a href="https://platform.openai.com/docs/api-reference/images">DALL-E</a> is then used to generate and download one (or more) <b>appropriate image(s)</b> to use within the presentation. Only tested on Linux 5.15.0-69 / Ubuntu 20.04.1 / Python 3.8.10. You will need to translate few GUI elements to your language (currently Italian) and customize paths. Code is adequately commented, with English prompts and instructions provided where necessary. 

<h3>Install, configure and run</h3>
<b>Download:</b> <code>git clone https://github.com/davidegat/G-PoinT.git</code>. Check Releases for compressed archives.<br>
<b>Python3 dependencies</b>: <code>pip install python-pptx requests Pillow openai</code><br>
<b>pptgui.py configuration</b>:<br>
<code>openai.api_key = "your_api_key"
template_path = "/path/to/your/template.pptx"
output_directory = "/path/to/your/output/directory"</code><br>
For further customization (temperature, prompts, tokens) look into code and change configuration variables.<br><br>
<b>Running and using G-PoinT</b>:<br>
<code>cd G-PoinT</code><br>
<code>python3 ./pptgui.py</code><br>
You can use the included G-PoinT.desktop file and access GUI via desktop, remember to edit paths accordingly and make it executable: <code>chmod +x G-PoinT.desktop</code> (see <a href="https://developer-old.gnome.org/desktop-entry-spec/">Desktop Entry Specifications</a>)<br><br>

Insert a topic and choose number of images to be generated, then click on generate button (for example, "Brain Tumor"), wait for a reasonable amount of time to receive PPTX and PNG outputs directly in custom folder. <b>Please note that it may take up to one minute to generate one PowerPoint and one image!</b> More images means more time.<br><br>
This repository contains <b>examples of output</b> and an <b>example template</b> you can customize or replace with your own.
<h3>Results from GPT and DALL-E</h3>
To obtain different results (more slides, more text, specific contexts), modify the <a href="https://help.openai.com/en/articles/6654000-best-practices-for-prompt-engineering-with-openai-api">prompt</a> at the beginning of the script directly. Try different prompts, temperatures, and tokens for fine-tuned results. If you need more realistic, artistic, or other styles for image generation, modify DALL-E prompt accordingly, read comments in code for details. Also refer to <a href="https://python-pptx.readthedocs.io/en/latest/">pptx library documentation</a> to customize font, colors, text size and other presentation elements, or if you want to include image directly into presentation.<br><br>
<b>Please note</b>: script actually works well generating 8 slides made of 6 key points each. If you need to increase slide number, amount of text, or to give GPT more "fantasy" editing the <a href="https://platform.openai.com/docs/api-reference/completions/create#completions/create-temperature">temperature</a> parameters, you should check for apropriate <a href="https://help.openai.com/en/articles/4936856-what-are-tokens-and-how-to-count-them">token size</a> an be ready to wait <b>longer generation times</b>, and pay more on API costs (still really low btw).<br><br>
<b>The following video shows process in real time (3 images generated).</b> Please wait until files are completely generated, or skip from 0:25 to 0:43.<br><br>

https://user-images.githubusercontent.com/51516281/230659957-7a52ab80-0148-4343-bf8c-cbadd85c8603.mp4

<h3>What it does?</h3>

<li>Shows Tkinter GUI and asks user for a topic and number of images to generate
<li>Sends prompt and topic to GPT
<li>Gets back 8 key points, each one will be a slide
<li>Sends one prompt for each key point to generate slide content text
<li>Creates a PPTX file from template with key points as slide titles, and fills them with generated text
<li>Translates topic from whatever language to english (also via GPT prompt)
<li>Asks DALL-E for images using topic in english
<li>Saves PPTX and PNG files.
<li>First slide will be empty for user customization.


<h3>Dependencies and requirements</h3>

<li>os - a built-in Python library for interacting with the operating system.
<li>tkinter - a built-in Python library for creating graphical user interfaces (GUIs).
<li>pptx - a Python library for creating and updating PowerPoint (.pptx) files.
<li>requests - a popular Python library for making HTTP requests.
<li>PIL (Python Imaging Library) - an open-source Python library for adding image processing capabilities to your Python interpreter.
<li>openai - the official Python library for the OpenAI API, used to interact with OpenAI services like GPT-3 and DALL-E.

Please note that the script assumes you have a Linux box and compatible version of Python 3 installed on your system (tested on Python 3.8.10). Additionally, script relies on having access to the <a href="https://platform.openai.com/account/api-keys">OpenAI API key</a>, which you'll need to $ign up for.
<h3>TO-DO</h3>
<li>Full English translation or multilanguage (maybe via GPT itself).
<li>Insert image(s) in PPTX directly.
<li>Count token used with each model, and inform user on costs for each generation.
<li>Speed up process, e.g. sending image request from the start in background, or parallel requests to API.
<li>Batch capabilities: insert multiple topics to batch generate presentations.
<li>Other platform (e.g. Windows) testing<br>
<br><b>Contributions and other ideas are welcome!</b>

<h3>Not professional. Not perfect.</h3>
G-PoinT is not a substitute for a real person in terms of content, but if you know what to do with prompts you could get good starting points to speed up your workflow, and get inspired. Try to play around with templates and pptx library for more impressive, yet not perfect, graphic results.
