# G-PoinT
![gpoint](https://user-images.githubusercontent.com/51516281/231153119-eb9f4c14-a269-449a-ad86-a5bfeac305a7.png)
<br><br>
This software uses Python3 and GPT via <a href="https://platform.openai.com/docs/api-reference/introduction">OpenAI API</a> to create a <b>complete</b> PowerPoint file in <b>any language</b>, including slides and text, <b>from a single topic input</b>. <a href="https://platform.openai.com/docs/api-reference/images">DALL-E</a> is then used to generate and download one (or more) <b>appropriate image(s)</b> to use within the presentation. Only tested on Linux 5.15.0-69 / Ubuntu 20.04.1 / Python 3.8.10. You will need to configure and customize paths. Code is adequately commented, with instructions provided where necessary. 

<h3>Install, configure and run</h3>
<b>Download:</b> <code>git clone https://github.com/davidegat/G-PoinT.git</code>. Check Releases for compressed archives.<br><br>
<b>Python3 dependencies</b>: <code>pip install python-pptx requests Pillow openai</code><br><br>
<b>Mandatory pptgui.py configuration</b>:<br>
<code>openai.api_key = "your_api_key"
template_path = "./templates/template.pptx"
template_folder = "./templates/"
output_directory = "/path/to/your/output/directory"</code><br>
<code>english = True</code> - For DALL-E prompt translation, set it to False if language is NOT in English.<br>
<code>language = "English"</code> - Set output language (any language understood by GPT).<br>
For further configurations, look variables and comments at the beginning of the script.<br><br>
<b>Running and using G-PoinT</b>:<br>
<code>cd G-PoinT</code><br>
<code>python3 ./pptgui.py</code><br>
You can use the included G-PoinT.desktop file and access GUI via desktop, remember to edit paths accordingly and make it executable: <code>chmod +x G-PoinT.desktop</code> (see <a href="https://developer-old.gnome.org/desktop-entry-spec/">Desktop Entry Specifications</a>)<br><br>

<li>Insert a topic (for example, "Dolphins", "General Relativity", "Heart Diseases")
<li>Insert number of pictures to be generated
<li>Choose picture size from menu
<li>Select your favourite template
<li>Click on button, wait for a reasonable amount of time to receive PPTX and PNG outputs directly in custom folder.<br><br><b>Please note that it may take up to one minute to generate one PowerPoint and one image!</b> More images means more time.<br><br>You can also choose to <b>only generate images</b> by clicking the "Generate more (or only) images" button.
This repository contains <b>examples of output</b> and <b>example templates</b> you can customize or replace with your own.
<h3>Results from GPT and DALL-E</h3>
For different results (more slides, more text, specific contexts), modify the <a href="https://help.openai.com/en/articles/6654000-best-practices-for-prompt-engineering-with-openai-api">prompt</a> at the beginning of the script directly. Try different prompts, temperatures, and tokens for fine-tuned results. If you need more realistic, artistic, or other styles for image generation, modify DALL-E prompt accordingly, read comments in code for details. Also refer to <a href="https://python-pptx.readthedocs.io/en/latest/">pptx library documentation</a> to customize font, colors, text size and other presentation elements.<br><br>
<b>Please note</b>: script actually works well generating 8 slides made of 6 points each. If you need to increase slide number, amount of text, or to give GPT more "fantasy" editing the <a href="https://platform.openai.com/docs/api-reference/completions/create#completions/create-temperature">temperature</a> parameters, you should check for apropriate <a href="https://help.openai.com/en/articles/4936856-what-are-tokens-and-how-to-count-them">token size</a> an be ready to wait <b>longer generation times</b>, and pay more on API costs (still really low btw).<br><br>

<h3>What it does?</h3>

<li>Shows Tkinter GUI and asks user for a topic and number of images to generate
<li>Sends prompt and topic to GPT
<li>Gets back 8 key points, each one will be a slide
<li>Sends one prompt for each key point to generate slide content text
<li>Creates a PPTX file from template with key points as slide titles, and fills them with generated text
<li>Translates topic from whatever language to English (also via GPT prompt - if not already in English).
<li>Asks DALL-E for images using topic in English
<li>Saves PPTX and PNG files.
<li>First slide will be empty for user customization

<h3>Dependencies and libraries explained</h3>

<li>`os`: Library to interact with the operating system.
<li>`requests`: Library to send HTTP/1.1 requests.
<li>`openai`: Library to interact with the OpenAI API.
<li>`glob`: Library to find files that match a specified pattern.
<li>`webbrowser`: Library to open a web browser from a Python script.
<li>`tkinter`: Library to create graphical user interfaces (GUIs) for desktop applications.
<li>`messagebox`: Module within `tkinter` to display message boxes and dialogs in a GUI application.
<li>`ttk`: Module within `tkinter` that provides a set of themed widgets.
<li>`BytesIO`: Library to manipulate binary data in memory as if it were a file.
<li>`pptx`: Library to create and manipulate Microsoft PowerPoint files.
<li>`Inches` and `Pt`: Modules within `pptx.util` that provide units of measurement for PowerPoint.
<li>`RGBColor`: Module within `pptx.dml.color` that provides a way to define colors for PowerPoint objects.

Please note that the script assumes you have a Linux box and compatible version of Python 3 installed on your system (tested on Python 3.8.10). Additionally, script relies on having access to the <a href="https://platform.openai.com/account/api-keys">OpenAI API key</a>, which you'll need to $ign up for.
<h3>TO-DO</h3>
<li>Count token used with each model, and inform user on costs for each generation.
<li>Speed up process, e.g. sending image request from the start in background, or parallel requests to API.
<li>Batch capabilities: insert multiple topics to batch generate presentations.
<li>Other platforms (e.g. Windows) testing<br>
<br><b>Contributions and ideas are welcome!</b>

<h3>Not professional. Not perfect.</h3>
G-PoinT is not a substitute for a real person in terms of contents, but if you know what to do with prompts you could get good starting points to speed up your workflow, and get inspired.
