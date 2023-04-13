# G-PoinT

![gpoint](https://user-images.githubusercontent.com/51516281/231254043-a65b5bee-75b5-4391-bb08-472becbda7f6.png)

G-PoinT is a software that interacts with GPT via [OpenAI API](https://platform.openai.com/docs/api-reference/introduction) to create a **complete** PowerPoint file in **any language**, including slides and text, **from a single topic input**. [DALL-E](https://platform.openai.com/docs/api-reference/images) is then used to generate and download one or more **picture(s)** to use within the presentation. G-PoinT can also generate a **presentation script** and **MP3 audio file** with the script reading. You will need to configure it with API KEY and custom paths in "config.ini" file. Code is adequately commented, with instructions provided where necessary.

## Tested on

- Linux 5.15.0-69, Ubuntu 20.04.1, Python 3.8.10.
- Windows 11, Python 3.11.3 (with pip enabled, and installed libraries), [Python for Windows](https://www.python.org/downloads/windows/) needed.

Please report any other working scenario!

## Help testing and developing

This code requires **lots** of API requests to be tested, maintained, upgraded, and hereby given for free. If you found it useful, please [consider supporting this project API costs with any small amount via PayPal](https://www.paypal.com/donate/?hosted_button_id=2EGA7T2LTD3AU). If you are a developer and want to contribute with **ideas and code**, you are welcome too!

Thanks for your sincere kindness! <3

## Examples

In this repository, you can find [examples in different languages](https://github.com/davidegat/G-PoinT/tree/main/examples): full presentations, presentation scripts, pictures, MP3 files. Take a look to see if this software can satisfy your needs! Files are uploaded as-is to understand both capabilities and limitations.

**Watch some videos of G-PoinT in action:**

- "Potato" slides, images, script, and audio **from first input to finished file**, example in English.
  https://user-images.githubusercontent.com/51516281/231296285-a56c027e-3f42-40c0-b988-5de4146e2fc5.mp4
- "Bonsai" slides in Italian:
  https://user-images.githubusercontent.com/51516281/231291250-5506a7dc-46ee-41e1-9ce7-402ef9022a7f.mp4
- "Tulips" slides in English:
  https://user-images.githubusercontent.com/51516281/231291271-30f5f722-65ed-427c-b817-904f4948e03f.mp4
- "Turing Test" slides in French:
  https://user-images.githubusercontent.com/51516281/231295684-448886f5-54c6-4501-b321-c28f774ec2ff.mp4
- "Albert Einstein" slides in English
  https://user-images.githubusercontent.com/51516281/231295977-36844280-51d5-49fe-b45b-fddc9c467007.mp4

## Install, configure, and run

Both for Windows and Linux machines, [git](https://git-scm.com/downloads) and [pip](https://pip.pypa.io/en/stable/installation/) must be installed. Type following commands in Linux Terminal or Windows PowerShell Terminal.

**Download:** `git clone https://github.com/davidegat/G-PoinT.git`\
or check [Releases](https://github.com/davidegat/G-PoinT/releases) for compressed archives.

**Python3 dependencies**: `pip install requests openai messagebox gTTS python-pptx`\
If you miss tkinter and glob for some reason: `pip install tkinter glob`

**Mandatory "config.ini" configuration**:\
`api_key = YOUR API KEY`\
`output_directory = Your desktop or favourite output folder`

Language can be changed "on the fly" via GUI (see [videos](https://github.com/davidegat/G-PoinT#examples)).

**Templates**:\
G-PoinT has some example templates working out of the box for testing, but you may want to replace them with your own. Just copy your favourite PowerPoint templates into "templates" folder before running G-PoinT. You will find them in the dropdown menu ready to be used. Please, do not delete the default template, or keep at least one file named "template.pptx" into the templates folder.

**Running and using G-PoinT**:\
Linux:
- Via terminal: `cd G-PoinT`
- `chmod +x ./pptgui.py` (type this command **only one time**)
- `./pptgui.py`
- You can use included G-PoinT.desktop file and access GUI via desktop, copy it to your desktop, remember to edit paths accordingly, and make it executable via terminal with: `chmod +x G-PoinT.desktop` (see [Desktop Entry Specifications](https://developer-old.gnome.org/desktop-entry-spec/)).

Windows:
- Open Windows PowerShell terminal (WIN+R, `powershell` -enter-)
- Move to G-PoinT folder and run pptgui.py
- `cd G-PoinT` (or folder where pptgui.py is)
- `python.exe pptgui.py`
- If you associated .py files with python.exe, this should work by opening pptgui.py directly.

**Usage:**
- Insert a topic (for example, "Dolphins", "General Relativity", "Heart Diseases")
- Insert number of pictures to be generated ("0" for no picture)
- Input picture size
- Select favourite template from drop-down menu
- Look at further generation options: you can generate both a **script** for your presentation and an **MP3 file** of it, to use within the presentation.
- Click on button, wait for a reasonable amount of time to receive PPTX and PNG outputs directly in custom folder.
- You can also choose to **only generate pictures** by clicking "I need only pictures" button.\
  **Please note that it may take up to one minute to generate one PowerPoint and one picture!** More pics mean more time.

**Language settings instructions**\
By modifying the "language" variable in config file, you will set your default language. Anyway, by clicking language menu, you can customize output to any language supported by GPT on the fly. To make it compatible with **gtts**, G-PoinT must obtain a language code from first characters of your input. Examples are:
- **it**alian
- **en**glish
- **de**utsch

Some languages may create chaos (e.g. Portuguese - pt, Chinese - zh...), but GPT can generate text only with language codes, in these cases **just input your language code** (**zh**, **pt**, **en**, **fi**...). You can use this format for **any** language if unsure.\
G-PoinT can work well with English input also if a different language is set, but best results are obtained if input is written in your own language (generates a better presentation script and MP3).

**Customize GPT results for text and pictures**\
For different results (more slides, more text, specific contexts), modify [prompt](https://help.openai.com/en/articles/6654000-best-practices-for-prompt-engineering-with-openai-api) in **config.ini**. Try different prompts, temperatures, and tokens for fine-tuned results. If you need more realistic, artistic, or other styles for picture generation, modify DALL-E prompt accordingly. Also, refer to [pptx library documentation](https://python-pptx.readthedocs.io/en/latest/) to customize font, colors, text size, and other presentation elements.

**Please note**: If you need to increase slide number, amount of text, or to give GPT more "fantasy" editing the [temperature](https://platform.openai.com/docs/api-reference/completions/create#completions/create-temperature) parameters, you should check for appropriate [token size](https://help.openai.com/en/articles/4936856-what-are-tokens-and-how-to-count-them) and be ready to wait **longer generation times** and pay more on API costs (still really low btw).

**What it does?**
- Shows Tkinter GUI and asks user for a topic and the number of pictures to generate
- Sends prompt and topic to GPT
- Gets back 8 key points, each one will be a slide
- Sends one prompt for each key point to generate slide content text
- Creates a PPTX file from template with key points as slide titles, and fills them with generated text
- Translates topic from whatever language to English (also via GPT prompt - if not already in English).
- Asks DALL-E for pictures using topic in English
- Saves PPTX and PNG files.
- Generates other files at user request: presentation script, MP3 audio of the script.
- The first slide will be empty for user customization

**Troubleshooting and Limitations**
1. Might *rarely* generate inappropriate or irrelevant content. In such cases, you can rerun the generation, or try with a modified prompt or temperature settings to achieve better results.
2. Ensure that you have a stable internet connection to use GPT and DALL-E APIs. Any interruption may cause G-PoinT to fail.
3. Might take some time to generate slides and images, depending on the number of slides, tokens, and images requested. Please be patient while G-PoinT works.
4. Increasing slides or content might lead to longer generation times and higher API costs.
5. Quality of images generated by DALL-E might not always meet expectations. you can rerun the generation, or try to adjust the DALL-E prompt and temperature settings to improve image results.
