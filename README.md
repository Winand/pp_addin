# PowerPoint Tools

This add-in for Microsoft PowerPoint includes several features for better productivity.

## Features

* **Embed data** - try to recover source data from charts and embed it into presentation file
* **Break links** - unlink selected chart from its source file
* **Clean designs** - remove all unused designs and templates from presentation file to reduce size
* **Send** - create new Outlook message and attach file with all or selected slides
* **Paste & Replace** - replace selected object on a slide with an object from the clipboard preserving position and ZOrder

## Installation

Releases include PPAM add-in files which can be loaded into PowerPoint directly.

`support.office` package is used for semi-automated builds but it's not yet available.

To build add-in file manually:
* Create new presentation
* Add modules listed in `manifest.py` file to VB project
* Save file as PPAM
