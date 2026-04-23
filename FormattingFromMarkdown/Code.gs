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
    .addItem('Rebuild Document from Markdown', 'rebuildFromMarkdown')
    .addToUi();
}

/**
 * Main execution function.
 */
function rebuildFromMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const rawText = body.getText();
  
  // Phase 1: Parse the raw text into structured blocks
  const { blocks, footnotes } = parseBlocks(rawText);
  
  // Phase 2: Clear the existing document
  body.clear();
  
  // Phase 3: Rebuild the document and apply styles in two passes
  buildDocument(doc, body, blocks, footnotes);
}

/**
 * Reads lines and categorizes them into block-level elements.
 */
function parseBlocks(rawText) {
  const lines = rawText.split('\n');
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
    if (line.trim().startsWith('|') && line.trim().endsWith('|')) {
      const rowData = line.split('|').slice(1, -1).map(c => c.trim());
      if (rowData.every(c => /^[-:]+$/.test(c))) continue; 

      const prev = blocks[blocks.length - 1];
      if (prev && prev.type === 'table') {
        prev.rows.push(rowData);
      } else {
        blocks.push({ type: 'table', rows: [rowData] });
      }
      continue;
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
 * Reconstructs the document body using parsed blocks via a 2-pass engine.
 */
function buildDocument(doc, body, blocks, footnotes) {
  const elementsToStyle = [];

  // PASS 1: STRUCTURE - Append RAW elements only
  blocks.forEach(block => {
    let element = null;

    if (block.type === 'heading') {
      const hEnum = DocumentApp.ParagraphHeading[`HEADING${block.level}`];
      element = body.appendParagraph(block.content);
      element.setHeading(hEnum);
      elementsToStyle.push({ el: element, block: block });
    } 
    else if (block.type === 'paragraph') {
      element = body.appendParagraph(block.content);
      elementsToStyle.push({ el: element, block: block });
    } 
    else if (block.type === 'code') {
      element = body.appendParagraph(block.content);
      elementsToStyle.push({ el: element, block: block });
    } 
    else if (block.type === 'table') {
      element = body.appendTable(block.rows);
      elementsToStyle.push({ el: element, block: block });
    } 
    else if (block.type === 'list') {
      // Just append the raw item. We will structure the tree in Pass 2.
      element = body.appendListItem(block.content);
      elementsToStyle.push({ el: element, block: block });
    }
  });

  // Remove empty top paragraph created by clear()
  if (body.getNumChildren() > 1 && body.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH && body.getChild(0).getText() === '') {
    body.getChild(0).removeFromParent();
  }

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
