# G-PoinT
Uses GPT to generate a PowerPoint file with slides and text, then uses DALL-E to generate one image to associate. Released AS-IS: you need to translate GPT prompts in your language (currently italian) and customize paths, code is commented.

<li>Built on Linux Ubuntu, Python 3, uses Tkinter GUI for topic input.
<li>Remember to customize with your API key and paths.
<li>Image will be placed in same path you chose for PPTX file.
<li>Needs a PowerPoint template file in the script folder (example included).
<li>First slide will be empty for user customization.

<h3>Python3 dependencies</h3>
<li>Pillow
<li>pptx
<li>requests
<li>openai
  
<h3>What it does?</h3>
<li>Shows Tkinter GUI and asks for topic
<li>Sends a prompt + topic to GPT
<li>Gets back 8 key points
<li>Sends another prompt to generate text for each slide
<li>Create a PPTX file with key points as slidew title, and filling them with generated text
<li>Translate the topic from whatever language to english (also via GPT prompt)
<li>Asks DALL-E for an image using topic in english
<li>Saves both PPTX and PNG file.

Not professional. Not perfect.
