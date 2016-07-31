Overview
========

This module simplifies the monotonous portions of writing a CTY evaluation by helping a savvy instructor to automate it. This module does not simplify the creative portion of writing in any way, it simply creates a template with each student's info and the standard opening paragraph filled in and programatically.

Usage
=====

Basics
------

To create a CTYEvaluation object call:

`myEval = CTYEvaluation(fname, lname, cname, date, course, instructor, site, ta, completion)`

The arguments, all of which are required are:
 * Student's legal first name
 * Student's legal last name
 * The common name you actually called the student (eg. Bill, not William). This argument can be None if the student went by their legal first name.
 * The date of the last Friday of the session
 * Official course title
 * Instructor's full name
 * Full site name (Hint: Any string will do, but CTY requires the format "Carlisle, PA".)
 * TA's full name
 * On of these phrases: "successfully completing", "completing", "participating in".
 
Then to save the file with the standard naming format, call:

`myEval.save()`

An optional suffix can be supplied as a string and will be appended to the CTY default name.

Signature Line
--------------
Adding the standard instructor's signature line is accomplished with:

`myEval.add_signature()`

Paragraphs
----------

While this module makes no attempt to autmoate the actual writing process, it does provide and API for those who may want to do so in some way. In order to add a body paragraph (or even just the paragraph header) to your eval, call:

`myEval.add_paragraph("Content Proficiency")`

or

`myEval.add_paragraph("Content Proficiency", "You did very well understanding the content...")`

Notes
-----

You can also add a TAs notes about a student to the eval in a bulleted list format for easier reference while writing.

```
notes = ['good at soldering', 'too talkative', 'likes transistors']
myEval.add_notes(notes)
```

This function supports but native python lists, and a custom xml-based format that I may document if anyone asks for it.


Dependancies
============
CTY requires evaluations to be submitted in Microsoft Open XML format and to that end, this module depends on the excellent python-docx module: https://github.com/python-openxml/python-docx

License
=======
This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or
distribute this software, either in source code form or as a compiled
binary, for any purpose, commercial or non-commercial, and by any
means.

In jurisdictions that recognize copyright laws, I dedicate any and all copyright interest in the software to the public domain. I make this dedication for the benefit of the public at large and to the detriment of my heirs and successors. I intend this dedication to be an overt act of relinquishment in perpetuity of all present and future rights to this software under copyright law.

For more information, please refer to <http://unlicense.org/>
