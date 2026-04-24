Google Apps Script Markdown tools
=================================

![screenshot of the Markdown Tools menu](./screenshot.png)

This addon provides two operations:
1. *Format Selected Markdown*: select some markdown in a google doc, and it will convert it to proper google docs formatting to the best of its abilities
2. *Markdownify Formatted Selection*: select some formatted text, and the tool will do a best effort replacement with equivalent markdown

Supports vanilla markdown, including:
- normal paragraphs, with inline **bold**, *italic*, `monospaced`, [links](example.com)
- code fences
- un/ordered lists, nested
- tables

The one obvious missing feature would be image blocks. Those get left alone.

Installation
------------

VBA macros style:
1. Click Extensions -> Apps Script
2. Copy the code from [Code.gs](./Code.gs) and paste it in the Apps Script editor for Code.gs
3. Hit save / CTRL+S
4. Reload the Doc
5. There is now a "Markdown Tools" menu.

At some point there will be a scareware screen about allowing "Unknown project" access to your Google Drive. The Unknown project is the thing you pasted, that is the project.

Alternatively, you can save this into a Google Apps Script on your drive and create a Test deployment targeting a google doc.
