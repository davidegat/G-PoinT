# G-PoinT
This program employs Python and GPT to create a complete PowerPoint presentation, including slides and text. Additionally, DALL-E is used to generate and download an appropriate image to accompany the presentation. Released AS-IS: you will need to translate GUI in your language (currently italian) and customize paths, code is adeguately commented with english prompts and instructions where needed.

<li>Built on Linux Ubuntu, Python 3, uses Tkinter GUI for topic input.
<li>Image will be placed in same path you choose for PPTX file.
<li>Needs a PowerPoint template file (example included).
<li>Remember to edit with your API key and paths.
  
<h3>What it does?</h3>
<li>Shows Tkinter GUI and asks for a topic
<li>Sends prompt + topic to GPT
<li>Gets back 8 key points, each one will be a slide
<li>Sends another prompt to generate text for each slide
<li>Creates a PPTX file with key points as slide titles, and fills them with generated text
<li>Translates topic from whatever language to english (also via GPT prompt)
<li>Asks DALL-E for an image using topic in english
<li>Saves both PPTX and PNG files.
<li>First slide will be empty for user customization.
<br><br>

Run GUI with:<br><br>
<code>python3 ./pptgui.py</code><br><br>
Insert a topic and push the generate button (e.g. "Brain Tumor"), wait a reasonable time to get PPTX and PNG outputs directly in custom folder.

<h3>Dependencies and requirements</h3>

<li>os - a built-in Python library for interacting with the operating system.
<li>tkinter - a built-in Python library for creating graphical user interfaces (GUIs).
<li>pptx - a Python library for creating and updating PowerPoint (.pptx) files.
<li>requests - a popular Python library for making HTTP requests.
<li>PIL (Python Imaging Library) - an open-source Python library for adding image processing capabilities to your Python interpreter.
<li>openai - the official Python library for the OpenAI API, used to interact with OpenAI services like GPT-3 and DALL-E.

To install the external dependencies, you can use pip:<br><br>
<code>pip install python-pptx requests Pillow openai</code><br><br>
Please note that the script assumes you have a compatible version of Python 3 (preferably Python 3.6 or later) installed on your system. Additionally, the script relies on having access to the OpenAI API key, which you'll need to sign up for.
<h3>Not professional. Not perfect.</h3>
