/*
BSD 2-Clause License

Copyright (c) 2026, Vlad Meșco

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

/**
 * onOpen hook to create the custom menu.
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Markdown Tools')
    .addItem('Format Selected Markdown', 'processSelectedMarkdown')
    .addItem('Markdownify Formatted Selection', 'markdownifySelectedText')
    .addToUi();
}

/**
 * Main execution function for selected text.
 */
function processSelectedMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (!selection) {
    DocumentApp.getUi().alert('Please highlight the markdown text you want to process first.');
    return;
  }

  const body = doc.getBody();
  let rawText = '';
  const rangeElements = selection.getRangeElements();
  let insertIndex = -1;
  const elementsToRemove = [];

  // Phase 1: Extract selected text and find the insertion point
  for (let i = 0; i < rangeElements.length; i++) {
    const rangeEl = rangeElements[i];
    const el = rangeEl.getElement();

    // Bubble up to find the direct child of the Body (e.g., Paragraph, ListItem)
    let topLevel = el;
    while (topLevel.getParent() && topLevel.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
      topLevel = topLevel.getParent();
    }

    // Capture the index of the first selected block
    if (insertIndex === -1) {
      insertIndex = body.getChildIndex(topLevel);
    }

    // Mark the block for deletion later
    if (!elementsToRemove.includes(topLevel)) {
      elementsToRemove.push(topLevel);
    }

    // Extract the raw text
    if (el.editAsText) {
      const textStr = el.asText().getText();
      if (rangeEl.isPartial()) {
        rawText += textStr.substring(rangeEl.getStartOffset(), rangeEl.getEndOffsetInclusive() + 1);
      } else {
        rawText += textStr;
      }
    }
    
    // Add a newline to separate blocks cleanly
    if (rangeEl.isPartial() || el.getType() === DocumentApp.ElementType.PARAGRAPH || el.getType() === DocumentApp.ElementType.LIST_ITEM) {
       rawText += '\n';
    }
  }
  
  // Phase 2: Parse the raw text into structured blocks
  const { blocks, footnotes } = parseBlocks(rawText);
  
  // Phase 3: Insert the parsed document blocks at the selection index
  insertDocumentBlocks(doc, body, insertIndex, blocks, footnotes);
  
  // Phase 4: Remove the originally selected blocks
  elementsToRemove.forEach(el => {
    try {
      el.removeFromParent();
    } catch(e) {
      // Fallback if it's the last paragraph in the document (which cannot be removed)
      if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
        el.clear();
      }
    }
  });
}

/**
 * REVERSE ENGINE: Converts selected Google Docs elements into Markdown
 */
function markdownifySelectedText() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (!selection) {
    DocumentApp.getUi().alert('Please highlight the text you want to Markdownify.');
    return;
  }

  const body = doc.getBody();
  const rangeElements = selection.getRangeElements();
  let insertIndex = -1;
  const elementsToProcess = [];

  // Phase 1: Identify full blocks involved in the selection
  const processedIndices = new Set(); // Track unique element indices

  for (let i = 0; i < rangeElements.length; i++) {
    const rangeEl = rangeElements[i];
    let topLevel = rangeEl.getElement();

    // Bubble up to find the direct child of the Body
    while (topLevel.getParent() && topLevel.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
      topLevel = topLevel.getParent();
    }

    const childIndex = body.getChildIndex(topLevel);

    if (insertIndex === -1) {
      insertIndex = childIndex;
    }

    // Deduplicate using the child index instead of the object reference
    if (!processedIndices.has(childIndex)) {
      console.log("Adding element to process:", topLevel.getType().toString(), "with text:", topLevel.getText ? topLevel.getText().substr(0, 40) : "[no text]");
      
      processedIndices.add(childIndex);
      elementsToProcess.push(topLevel);
    }
  }

  // Phase 2: Traverse elements and generate Markdown strings
  let mdLines = [];
  let footnotes = [];
  let inCodeFence = false;

  elementsToProcess.forEach(el => {
    const type = el.getType();
    console.log("Processing element type:", type.toString(), "with text:", el.getText ? el.getText().substr(0, 40) : "[no text]");

    // Handle Code Fences
    if (type === DocumentApp.ElementType.PARAGRAPH && isMonospaceBlock(el)) {
      if (!inCodeFence) {
        mdLines.push('```');
        inCodeFence = true;
      }
      mdLines.push(el.getText()); // Raw text, no inline markdown processing
      return;
    } else {
      if (inCodeFence) {
        mdLines.push('```');
        inCodeFence = false;
      }
    }

    // Process different block types
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const headingMap = {
        [DocumentApp.ParagraphHeading.HEADING1]: '# ',
        [DocumentApp.ParagraphHeading.HEADING2]: '## ',
        [DocumentApp.ParagraphHeading.HEADING3]: '### ',
        [DocumentApp.ParagraphHeading.HEADING4]: '#### ',
        [DocumentApp.ParagraphHeading.HEADING5]: '##### ',
        [DocumentApp.ParagraphHeading.HEADING6]: '###### '
      };
      const prefix = headingMap[el.getHeading()] || '';
      const text = extractMarkdownFromContainer(el, footnotes);
      if (text.trim() !== '' || prefix !== '') {
        mdLines.push(prefix + text);
      } else {
        mdLines.push('');
      }
    } 
    else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const indent = '  '.repeat(el.getNestingLevel());
      const orderedGlyphs = [
        DocumentApp.GlyphType.NUMBER,
        DocumentApp.GlyphType.LATIN_LOWER,
        DocumentApp.GlyphType.LATIN_UPPER,
        DocumentApp.GlyphType.ROMAN_LOWER,
        DocumentApp.GlyphType.ROMAN_UPPER
      ];
      const unorderedChars = [
        '-',
        '+',
        '*'
      ];
      const isOrdered = orderedGlyphs.includes(el.getGlyphType());
      const prefix = isOrdered ? '1. ' : `${unorderedChars[el.getNestingLevel() % unorderedChars.length]} `;
      mdLines.push(indent + prefix + extractMarkdownFromContainer(el, footnotes));
    } 
    else if (type === DocumentApp.ElementType.TABLE) {
      console.log("Processing table with", el.getNumRows(), "rows");
      for (let r = 0; r < el.getNumRows(); r++) {
        const row = el.getRow(r);
        console.log("Processing row", r, "with", row.getNumCells(), "cells");
        let rowData = [];
        for (let c = 0; c < row.getNumCells(); c++) {
          console.log("Processing cell", c);
          let cell = row.getCell(c);
          let cellMdLines = [];
          for (let p = 0; p < cell.getNumChildren(); p++) {
             cellMdLines.push(extractMarkdownFromContainer(cell.getChild(p), footnotes));
          }
          console.log("Add row", r, "cell", c, "content:", cellMdLines.join(' | '));
          // Flatten multi-paragraph cells to `<br>` for markdown tables
          rowData.push(cellMdLines.join('<br>')); 
        }
        mdLines.push('| ' + rowData.join(' | ') + ' |');
        
        // Add table header separator
        if (r === 0) {
          mdLines.push('|' + rowData.map(() => '---').join('|') + '|');
        }
      }
    }
  });

  if (inCodeFence) mdLines.push('```');

  // Append Footnotes
  if (footnotes.length > 0) {
    mdLines.push('');
    footnotes.forEach((note, index) => {
      mdLines.push(`[^${index + 1}]: ${note}`);
    });
  }

  // Phase 3: Insert raw text and remove original objects
  const finalMarkdown = mdLines.join('\n');
  body.insertParagraph(insertIndex, finalMarkdown);

  elementsToProcess.forEach(el => {
    try {
      el.removeFromParent();
    } catch(e) {
      if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
        el.clear();
      }
    }
  });
}

/**
 * Extracts inner contents and transforms inline Google Docs styles to Markdown
 */
function extractMarkdownFromContainer(container, footnotes) {
  let mdText = '';
  // Check if standard container with children
  if (typeof container.getNumChildren !== 'function') return container.getText();

  for (let i = 0; i < container.getNumChildren(); i++) {
    const child = container.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.TEXT) {
      mdText += processTextRun(child);
    } else if (type === DocumentApp.ElementType.FOOTNOTE) {
      const fnContents = child.getFootnoteContents().getText().trim();
      footnotes.push(fnContents);
      mdText += `[^${footnotes.length}]`;
    } else {
      // Best effort for unhandled inline items (like equations)
      if (child.getText) mdText += child.getText();
    }
  }
  return mdText;
}

/**
 * Helper to extract and format attributes in an inline Text Element
 */
function processTextRun(textEl) {
  const rawText = textEl.getText();
  if (!rawText) return '';
  
  const indices = textEl.getTextAttributeIndices();
  let mdStr = '';

  for (let i = 0; i < indices.length; i++) {
    const start = indices[i];
    const end = i + 1 < indices.length ? indices[i + 1] - 1 : rawText.length - 1;
    let seg = rawText.substring(start, end + 1);

    // Pull out whitespace to prevent markdown syntax breaking, e.g., "** bold **"
    const leadingSpace = seg.match(/^\s*/)[0];
    const trailingSpace = seg.match(/\s*$/)[0];
    let coreSeg = seg.trim();

    if (!coreSeg) {
      mdStr += seg; 
      continue;
    }

    const font = textEl.getFontFamily(start);
    const isCode = ['Courier New', 'Consolas', 'Monaco', 'monospace'].includes(font);
    const isBold = textEl.isBold(start);
    const isItalic = textEl.isItalic(start);
    const isUnderline = textEl.isUnderline(start);
    const linkUrl = textEl.getLinkUrl(start);

    // Apply inline syntax (inside out)
    if (isCode) coreSeg = `\`${coreSeg}\``;
    if (isItalic) coreSeg = `*${coreSeg}*`;
    if (isBold) coreSeg = `**${coreSeg}**`;
    if (isUnderline) coreSeg = `<u>${coreSeg}</u>`;
    if (linkUrl) coreSeg = `[${coreSeg}](${linkUrl})`;

    mdStr += leadingSpace + coreSeg + trailingSpace;
  }
  
  return mdStr;
}

/**
 * Helper: Detects if an entire paragraph is uniformly styled as monospace.
 */
function isMonospaceBlock(container) {
  if (container.getType() !== DocumentApp.ElementType.PARAGRAPH) return false;
  if (container.getNumChildren() === 0) return false;
  
  let hasText = false;
  for (let i = 0; i < container.getNumChildren(); i++) {
    const child = container.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TEXT) {
      hasText = true;
      const font = child.getFontFamily(0);
      if (!['Courier New', 'Consolas', 'Monaco', 'monospace'].includes(font)) {
        return false;
      }
    } else if (child.getType() !== DocumentApp.ElementType.FOOTNOTE) {
      return false; // Inline images, etc., break the "pure block" assumption
    }
  }
  return hasText;
}

// ---------------------------------------------------------
// ORIGINAL PARSING FUNCTIONS BELOW
// ---------------------------------------------------------

/**
 * Reads lines and categorizes them into block-level elements.
 */
function parseBlocks(rawText) {
  const lines = rawText.replaceAll('\r\n', '\n').replaceAll('\r', '\n').split('\n');
  const blocks = [];
  const footnotes = {};
  
  let inCodeFence = false;
  let activeFence = '';
  let codeContent = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // 1. Code Fences
    const fenceMatch = line.trim().match(/^(`{3,})/);
    if (fenceMatch) {
      const matchedFence = fenceMatch[1];
      
      if (!inCodeFence) {
        inCodeFence = true;
        activeFence = matchedFence;
        continue;
      } else if (inCodeFence && matchedFence.length >= activeFence.length) {
        blocks.push({ type: 'code', content: codeContent.join('\n') });
        inCodeFence = false;
        activeFence = '';
        codeContent = [];
        continue;
      }
    }

    if (inCodeFence) {
      codeContent.push(line);
      continue;
    }

    // 2. Footnote Definitions: [^id]: text
    const fnMatch = line.match(/^\[\^([^\]]+)\]:\s*(.*)/);
    if (fnMatch) {
      footnotes[fnMatch[1]] = fnMatch[2];
      continue;
    }

    // 3. Headings
    const headingMatch = line.match(/^(#{1,6})\s+(.*)/);
    if (headingMatch) {
      blocks.push({ type: 'heading', level: headingMatch[1].length, content: headingMatch[2] });
      continue;
    }

    // 4. Tables
    {
      const prev = blocks[blocks.length - 1];
      if (line.trim().startsWith('|') && (line.trim().endsWith('|') || (prev && prev.type === 'table'))) {
        const rowData = line.split('|').slice(1, line.trim().endsWith('|') ? -1 : undefined).map(c => c.trim());
        if (rowData.every(c => /^[-:]+$/.test(c))) continue; 

        if (prev && prev.type === 'table') {
          console.log("Appending row to existing table:", rowData);
          prev.rows.push(rowData);
        } else {
          console.log("starting table", rowData);
          console.log("line:", `{{${line}}}`, line.split('').map(c => c.charCodeAt(0)));
          blocks.push({ type: 'table', rows: [rowData] });
        }
        continue;
      }
    }

    // 5. Lists (Ordered and Unordered, with indentation detection)
    const listMatch = line.match(/^(\s*)([-*+]|\d+\.)\s+(.*)/);
    if (listMatch) {
      const indent = listMatch[1].length;
      const isOrdered = /\d+\./.test(listMatch[2]);
      blocks.push({ 
        type: 'list', 
        level: Math.floor(indent / 2), 
        isOrdered: isOrdered, 
        content: listMatch[3] 
      });
      continue;
    }

    // 6. Paragraphs
    if (line.trim() !== '') {
      blocks.push({ type: 'paragraph', content: line });
    } else {
      blocks.push({ type: 'paragraph', content: '' });
    }
  }

  return { blocks, footnotes };
}

/**
 * Reconstructs the document body by inserting parsed blocks at a specific index via a 2-pass engine.
 */
function insertDocumentBlocks(doc, body, startIndex, blocks, footnotes) {
  const elementsToStyle = [];
  let currentIndex = startIndex;

  // PASS 1: STRUCTURE - Insert RAW elements at the specific index
  blocks.forEach(block => {
    let element = null;

    if (block.type === 'heading') {
      const hEnum = DocumentApp.ParagraphHeading[`HEADING${block.level}`];
      element = body.insertParagraph(currentIndex, block.content);
      element.setHeading(hEnum);
      elementsToStyle.push({ el: element, block: block });
      currentIndex++;
    } 
    else if (block.type === 'paragraph') {
      element = body.insertParagraph(currentIndex, block.content);
      elementsToStyle.push({ el: element, block: block });
      currentIndex++;
    } 
    else if (block.type === 'code') {
      element = body.insertParagraph(currentIndex, block.content);
      elementsToStyle.push({ el: element, block: block });
      currentIndex++;
    } 
    else if (block.type === 'table') {
      element = body.insertTable(currentIndex, block.rows);
      elementsToStyle.push({ el: element, block: block });
      currentIndex++;
    } 
    else if (block.type === 'list') {
      element = body.insertListItem(currentIndex, block.content);
      elementsToStyle.push({ el: element, block: block });
      currentIndex++;
    }
  });

  // PASS 2: PAINT AND BIND
  let currentListParent = null; 

  elementsToStyle.forEach(item => {
    const { el, block } = item;
    
    // Break the list chain if we encounter a non-list element
    if (block.type !== 'list') {
      currentListParent = null;
    }

    if (block.type === 'code') {
      el.setFontFamily('Courier New');
      el.setBackgroundColor('#f3f4f6');
    } 
    else if (block.type === 'table') {
      for (let r = 0; r < el.getNumRows(); r++) {
        for (let c = 0; c < el.getRow(r).getNumCells(); c++) {
          processInlineStyles(doc, el.getRow(r).getCell(c).editAsText(), footnotes);
        }
      }
    } 
    else {
      // Process inline styles FIRST (this mutates the string and breaks index 0)
      if (block.content !== '') {
        processInlineStyles(doc, el.editAsText(), footnotes);
      }

      // NOW that the text string is final, apply list structure and glyphs safely
      if (block.type === 'list') {
        if (currentListParent) {
          el.setListId(currentListParent);
        } else {
          currentListParent = el;
        }
        
        el.setNestingLevel(block.level);
        
        const levelDepth = block.level % 3;
        if (block.isOrdered) {
          const orderedGlyphs = [
            DocumentApp.GlyphType.NUMBER,
            DocumentApp.GlyphType.LATIN_LOWER,
            DocumentApp.GlyphType.ROMAN_LOWER
          ];
          el.setGlyphType(orderedGlyphs[levelDepth]);
        } else {
          const unorderedGlyphs = [
            DocumentApp.GlyphType.BULLET,
            DocumentApp.GlyphType.HOLLOW_BULLET,
            DocumentApp.GlyphType.SQUARE_BULLET
          ];
          el.setGlyphType(unorderedGlyphs[levelDepth]);
        }
      }
    }
  });
}

/**
 * Executes the inline formatting pipeline on a specific Text element.
 */
function processInlineStyles(doc, textElement, footnotes) {
  applyInlineStyle(textElement, '\\*\\*([^\\*]+)\\*\\*', 2, 2, (start, end) => textElement.setBold(start, end, true));
  applyInlineStyle(textElement, '\\*([^\\*]+)\\*', 1, 1, (start, end) => textElement.setItalic(start, end, true));
  applyInlineStyle(textElement, '<u>(.*?)<\\/u>', 3, 4, (start, end) => textElement.setUnderline(start, end, true));
  applyInlineStyle(textElement, '`([^`]+)`', 1, 1, (start, end) => {
    textElement.setFontFamily(start, end, 'Courier New');
    textElement.setBackgroundColor(start, end, '#f3f4f6');
  });
  processLinks(textElement);
  processFootnotes(doc, textElement, footnotes);
}

/**
 * Generic inline styling helper for matched regex.
 */
function applyInlineStyle(textElement, regexStr, leftMarkerLen, rightMarkerLen, styleCallback) {
  let found = textElement.findText(regexStr);
  while (found) {
    let start = found.getStartOffset();
    let end = found.getEndOffsetInclusive();
    styleCallback(start + leftMarkerLen, end - rightMarkerLen);
    textElement.deleteText(end - rightMarkerLen + 1, end);
    textElement.deleteText(start, start + leftMarkerLen - 1);
    found = textElement.findText(regexStr);
  }
}

/**
 * Extracts URL data and applies the link style.
 */
function processLinks(textElement) {
  const linkRegex = '\\[([^\\]]+)\\]\\(([^)]+)\\)';
  let found = textElement.findText(linkRegex);
  while (found) {
    let start = found.getStartOffset();
    let end = found.getEndOffsetInclusive();
    let fullText = textElement.getText().substring(start, end + 1);
    let match = fullText.match(/\[([^\]]+)\]\(([^)]+)\)/);
    if (match) {
      let display = match[1];
      let url = match[2];
      textElement.deleteText(start, end);
      textElement.insertText(start, display);
      textElement.setLinkUrl(start, start + display.length - 1, url);
    }
    found = textElement.findText(linkRegex);
  }
}

/**
 * Replaces footnote markers with native Google Docs footnotes.
 */
function processFootnotes(doc, textElement, footnotesMap) {
  const fnRegex = '\\[\\^([^\\]]+)\\]';
  let found = textElement.findText(fnRegex);
  while (found) {
    let start = found.getStartOffset();
    let end = found.getEndOffsetInclusive();
    let fullText = textElement.getText().substring(start, end + 1);
    let match = fullText.match(/\[\^([^\]]+)\]/);
    if (match) {
      let id = match[1];
      let noteText = footnotesMap[id] || "Missing footnote definition";
      textElement.deleteText(start, end);
      try {
        let pos = doc.newPosition(textElement, start);
        let fn = doc.addFootnote(pos);
        fn.getFootnoteContents().getParagraphs()[0].setText(noteText);
      } catch(e) {
        textElement.insertText(start, ` [Footnote: ${noteText}]`);
      }
    }
    found = textElement.findText(fnRegex);
  }
}
