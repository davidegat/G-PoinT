# G-PoinT
![gpoint](https://user-images.githubusercontent.com/51516281/231254043-a65b5bee-75b5-4391-bb08-472becbda7f6.png)
<br><br>
This software interacts with GPT via <a href="https://platform.openai.com/docs/api-reference/introduction">OpenAI API</a> to create a <b>complete</b> PowerPoint file in <b>any language</b>, including slides and text, <b>from a single topic input</b>. <a href="https://platform.openai.com/docs/api-reference/images">DALL-E</a> is then used to generate and download one or more <b>picture(s)</b> to use within the presentation. G-PoinT can also generate a <b>presentation script</b> and <b>MP3 audio file</b> with the script reading.<br><br>
Tested on:
<li>Linux 5.15.0-69, Ubuntu 20.04.1, Python 3.8.10.
<li>Windows 11, Python 3.11.3 (with pip enabled, and installed libraries), <a href="https://www.python.org/downloads/windows/">Python for Windows</a> needed. <br><br>You will need just to configure it with API KEY and customize paths in. Please report any working scenario to upgrade this list! Code is adequately commented, with instructions provided where necessary. 
<h3>Help testing and developing</h3>
This code requires <b>lots</b> of API requests to be tested, mantained, upgraded, and hereby given for free. If you found it useful, please <a href="https://www.paypal.com/donate/?hosted_button_id=2EGA7T2LTD3AU">consider supporting this project API costs with any small amount via PayPal</a>. <br>If you are a developer and want to contribute with <b>ideas and code</b>, you are welcome too!<br><br>Thanks for your sincere kindness! <3<br>

<h3>Examples</h3>
In this repository you can find <a href="https://github.com/davidegat/G-PoinT/tree/main/examples">examples in different languages</a>: full presentations, presentation scripts, pictures, MP3 files. Take a look to see if this software can satisfy your needs! Files are uploaded as-is to understand both capabilities and limitations.<br><br>

<b>Watch some videos of G-PoinT in action:</b><br>
<li>"Potato" slides, images, script and audio <b>from first input to finished file</b>, example in English.<br>
https://user-images.githubusercontent.com/51516281/231296285-a56c027e-3f42-40c0-b988-5de4146e2fc5.mp4
<li>"Bonsai" slides in Italian:<br>
https://user-images.githubusercontent.com/51516281/231291250-5506a7dc-46ee-41e1-9ce7-402ef9022a7f.mp4
<li>"Tulips" slides in English:<br>
https://user-images.githubusercontent.com/51516281/231291271-30f5f722-65ed-427c-b817-904f4948e03f.mp4
<li>"Turing Test" slides in French:<br>
https://user-images.githubusercontent.com/51516281/231295684-448886f5-54c6-4501-b321-c28f774ec2ff.mp4
<li>"Albert Einstein" slides in English<br>
https://user-images.githubusercontent.com/51516281/231295977-36844280-51d5-49fe-b45b-fddc9c467007.mp4

<h3>Install, configure and run</h3>
<b>Download:</b> <code>git clone https://github.com/davidegat/G-PoinT.git</code>.<br>Check <a href="https://github.com/davidegat/G-PoinT/releases">Releases</a> for compressed archives.<br><br>
<b>Python3 dependencies</b>: <code>pip install requests openai messagebox gTTS python-pptx</code><br>
If you miss for some reason tkinter and glob: <code>pip install tkinter glob</code><br><br>
<b>Mandatory pptgui.py configuration</b>:<br>
<code>openai.api_key = "your_api_key"
template_path = "./templates/template.pptx"
template_folder = "./templates/"
output_directory = "/path/to/your/output/directory"</code><br>
For further configurations, look variables and comments at the beginning of the script. Language can also be changed "on the fly" via GUI (see <a href="https://github.com/davidegat/G-PoinT#examples">videos</a>).<br><br>
<b>Templates</b>:<br>
G-PoinT has some example templates working out of the box for testing, but you may want to replace them with your own. Just copy your favourite PowerPoint templates into "templates" folder before running G-PoinT. You will find them in dropdown menu ready to be used. Please, do not delete default template, or keep at least one file named "template.pptx" into the templates folder.<br><br>
<b>Running and using G-PoinT</b>:<br>
Via Linux or Windows terminal type:<br>
<code>cd G-PoinT</code><br> (both Linux and Windows)
<code>chmod +x ./pptgui.py</code> (Linux only - type this command <b>only one time</b>)<br>
<code>./pptgui.py</code><br> (Windows: <code>python.exe pptgui.py</code><br>
For Linux: you can use the included G-PoinT.desktop file and access GUI via desktop, copy it to your desktop, remember to edit paths accordingly, and make it executable via terminal with: <code>chmod +x G-PoinT.desktop</code> (see <a href="https://developer-old.gnome.org/desktop-entry-spec/">Desktop Entry Specifications</a>).<br><br>
<li>Insert a topic (for example, "Dolphins", "General Relativity", "Heart Diseases")
<li>Insert number of pictures to be generated ("0" for no picture)
<li>Input picture size
<li>Select favourite template from drop down menu
<li>Look at further generation options: you can generate both a <b>script</b> for your presentation and an <b>MP3 file</b> of it, to use within the presentation.
<li>Click on button, wait for a reasonable amount of time to receive PPTX and PNG outputs directly in custom folder.
<li>You can also choose to <b>only generate pictures</b> by clicking the "I need only pictures" button.<br><br><b>Please note that it may take up to one minute to generate one PowerPoint and one picture!</b> More pics means more time.
<h3>Language settings instructions</h3>
By modifying the "language" variable into the script, you will set your default language. Anyway, by clicking the language menu, you can customize output to any language supported by GPT on the fly. To make it compatible with <b>gtts</b>, G-PoinT must obtain a language code from first characters of your input. Examples are:<br>
<li><b>it</b>alian
<li><b>en</b>glish
<li><b>de</b>utsch<br><br>
Some languages may create chaos (e.g. Portuguese - pt, Chinese - zh...), but GPT can generate text only with language codes, in these cases <b>just input your language code</b> (<b>zh</b>, <b>pt</b>, <b>en</b>, <b>fi</b>..). You can use this format for <b>any</b> language if unsure.<br><br>G-PoinT can work well with English input also if a different language is set, but best results are obtained if input is written in your own language (generates a better presentation script and MP3).
<h3>Customize GPT results for text and pictures</h3>
For different results (more slides, more text, specific contexts), modify the <a href="https://help.openai.com/en/articles/6654000-best-practices-for-prompt-engineering-with-openai-api">prompt</a> at the beginning of the script directly. Try different prompts, temperatures, and tokens for fine-tuned results. If you need more realistic, artistic, or other styles for picture generation, modify DALL-E prompt accordingly, read comments in code for details. Also refer to <a href="https://python-pptx.readthedocs.io/en/latest/">pptx library documentation</a> to customize font, colors, text size and other presentation elements.<br><br>
<b>Please note</b>: script actually works well generating 8 slides made of 6 points each. If you need to increase slide number, amount of text, or to give GPT more "fantasy" editing the <a href="https://platform.openai.com/docs/api-reference/completions/create#completions/create-temperature">temperature</a> parameters, you should check for apropriate <a href="https://help.openai.com/en/articles/4936856-what-are-tokens-and-how-to-count-them">token size</a> an be ready to wait <b>longer generation times</b>, and pay more on API costs (still really low btw).

<h3>What it does?</h3>

<li>Shows Tkinter GUI and asks user for a topic and number of pictures to generate
<li>Sends prompt and topic to GPT
<li>Gets back 8 key points, each one will be a slide
<li>Sends one prompt for each key point to generate slide content text
<li>Creates a PPTX file from template with key points as slide titles, and fills them with generated text
<li>Translates topic from whatever language to English (also via GPT prompt - if not already in English).
<li>Asks DALL-E for pictures using topic in English
<li>Saves PPTX and PNG files.
<li>Generates other files at user request: presentation script, MP3 audio of the script.
<li>First slide will be empty for user customization

<h3>Dependencies and libraries explained</h3>
<code>pip install requests openai messagebox gTTS python-pptx</code><br>
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
<li>`gTTS`: A library for converting text to speech.

Please note that the script assumes you have a Linux box and compatible version of Python 3 installed on your system (tested on Python 3.8.10). Additionally, script relies on having access to the <a href="https://platform.openai.com/account/api-keys">OpenAI API key</a>, which you'll need to $ign up for.
<h3>TO-DO</h3>
<li>Count token used with each model, and inform user on costs for each generation.
<li>Speed up process, e.g. sending picture request from the start in background, or parallel requests to API.
<li>Batch capabilities: insert multiple topics to batch generate presentations.
<li>Other platforms (e.g. Windows) testing<br>
<br><b>Contributions and ideas are welcome!</b>

<h3>Not professional. Not perfect.</h3>
G-PoinT is not a substitute for a real person in terms of contents, but if you know what to do with prompts you could get good starting points to speed up your workflow, and get inspired.
