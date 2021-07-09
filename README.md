# vtpy
A helper library for working with VTScada. Still very much in development, but can be quite useful for writing quick scripts. I wrote most of this on an as-needed basis, so it's far from perfect and may require looking at the code to understand everything that can be done, unfortunately. Use at your own risk, and be sure to check any data before putting it into a database.  
I did comment a fair amount though, so it shouldn't be too hard to figure out. I'm especially sorry about the database connection management, but it works fine for running quick scripts, so I haven't fixed it just yet.

This module requires that you export the tag database to an Access DB (.mdb).

To have this library available for import in any script, simply create a folder in your site-packages directory called "vtpy", and place "\_\_init\_\_.py" into that folder.
You can find the site-packages directory by entering "where python" in the command prompt, and replacing ```"\python.exe"``` with ```"\Lib\site-packages"```.
For example, on my machine it's located at ```"C:\Users\[USER]\AppData\Local\Programs\Python\Python39\Lib\site-packages"```

A bit of example script to get you started:
```python3
from vtpy import Tag, DBConnection

db_path = r'C:\EXPORTDB.mdb'
db = DBConnection(v_path)
tags = v_db.get_tags()

for tag in tags:
  read_address = tag.get("ReadAddress") # Get any 
  print(tag) # outputs a tab-separated string that you could copy and paste into Access or Excel
  # then you can do whatever you like with that
```
There's a bunch of other potentially useful functions in Tag and DBConnection too, so be sure to check that out.

----------------- 
  
  
  
  
I've also added some helper functions you can import such as:

```ParseIFixCSV``` which returns a list of dictionaries of ```{column_name: value}```. Each list entry is a tag, and the dictionary is its attributes. Example attributes might be things like "TAG", "DESCRIPTION", "I/O DEVICE", etc (case-sensitive). This can be nice for helping to convert tags between Fix and VT, for example.

```GetPages``` which returns a dictionary of ```{page_name: page_contents_as_text}``` which, for example, I've found useful when converting widgets on a lot of pages. I was able to use this to get the text of each page, find-and-replace (very carefully!) the widget I wanted to change, then output to files which I then imported into VT (again, making sure to check each change as I imported).

```GetTagValues``` which returns a dictionary of ```{tag_id: {tag_property: value}}``` (yes, a dict of dicts) which I've used to get the actual values for things like ReadAddress when the database only has an expression. If used right, this *should* give you access to the value that each expression is currently evaluated to.  
Annoyingly, the IDs don't line up one-to-one with those in the DB, so you'll need to use something like this to convert:  
```python3
tag.get(Tag.id_col).split(',')[0] if ',' in tag.get(Tag.id_col) else tag.get(Tag.id_col)
```
(also, what are they doing storing this data in text files anyway?)

----------

Lastly, I want to give a shout-out to pyperclip, which I use even more than this library for quick scripts. Once installed (```pip install pyperclip```) and imported, simply use ```pyperclip.paste()``` and ```pyperclip.copy(text)``` to work with text data (from Excel, VT Idea Studio, etc) without ever needing to write to a file or manually copy console output.
