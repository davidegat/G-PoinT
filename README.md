# G-PoinT
Uses Python and GPT to generate a full PowerPoint file with slides and text, and DALL-E to get one image to associate. Released AS-IS: you need to translate some GUI parts in your language (currently italian) and customize paths, code is commented with english prompts and instructions where needed.

<li>Built on Linux Ubuntu, Python 3, uses Tkinter GUI for topic input.
<li>Image will be placed in same path you choose for PPTX file.
<li>Needs a PowerPoint template file in the script folder (example included).
<li>First slide will be empty for user customization.
<li>Remember to edit with your API key and paths.

<h3>Python3 dependencies</h3>
<li>Pillow
<li>pptx
<li>requests
<li>openai
  
<h3>What it does?</h3>
<li>Shows Tkinter GUI and asks for a topic
<li>Sends prompt + topic to GPT
<li>Gets back 8 key points, each one will be a slide
<li>Sends another prompt to generate text for each slide
<li>Creates a PPTX file with key points as slide titles, and fills them with generated text
<li>Translates topic from whatever language to english (also via GPT prompt)
<li>Asks DALL-E for an image using topic in english
<li>Saves both PPTX and PNG files.

Not professional. Not perfect.
