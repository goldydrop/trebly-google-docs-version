function onOpen() {
  DocumentApp.getUi()
    .createMenu('Screenplay Editor')
    .addItem('Launch Trelby Editor', 'showTrelbyModal')
    .addToUi();
}

function showTrelbyModal() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setWidth(1150)
      .setHeight(780);
  DocumentApp.getUi().showModalDialog(html, 'Trelby Screenplay Mode');
}

// Reads existing paragraphs from the active Google Doc and detects their Trelby element types
function loadScriptFromDoc() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  const elements = [];

  paragraphs.forEach((p) => {
    const text = p.getText();
    if (!text || text.trim() === '') return;

    const indent = p.getIndentStart() || 0;
    const align = p.getAlignment();
    const isBold = p.editAsText().isBold(0) === true;

    let type = 'Action';

    if (align === DocumentApp.HorizontalAlignment.RIGHT) {
      type = 'Transition';
    } else if (indent >= 165) {
      type = 'Character';
    } else if (indent >= 125) {
      type = 'Parenthetical';
    } else if (indent >= 70) {
      type = 'Dialogue';
    } else if (isBold || /^(INT\.|EXT\.|INT\/EXT\.|I\/E\.|\d+\.\s*(INT\.|EXT\.))/i.test(text.trim())) {
      type = 'Scene';
    }

    elements.push({
      type: type,
      text: text
    });
  });

  return {
    title: doc.getName(),
    elements: elements
  };
}

function syncFullScript(elements) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const oldChildCount = body.getNumChildren();

  const validElements = elements.filter(item => item.text && item.text.trim() !== '');

  if (validElements.length === 0) {
    body.appendParagraph('');
    for (let i = oldChildCount - 1; i >= 0; i--) {
      try { body.removeChild(body.getChild(i)); } catch (e) {}
    }
    doc.saveAndClose();
    return 'Cleared';
  }

  validElements.forEach((item) => {
    const text = item.text.trim();
    const p = body.appendParagraph(text);

    p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    p.setFontFamily('Courier Prime');
    p.setFontSize(12);
    p.editAsText().setBold(false);

    switch (item.type) {
      case 'Scene':
        p.setText(text.toUpperCase());
        p.editAsText().setBold(true);
        p.setIndentFirstLine(0);
        p.setIndentStart(0);
        p.setIndentEnd(0);
        p.setSpacingBefore(12);
        p.setSpacingAfter(12);
        break;
      case 'Action':
        p.setIndentFirstLine(0);
        p.setIndentStart(0);
        p.setIndentEnd(0);
        p.setSpacingBefore(0);
        p.setSpacingAfter(12);
        break;
      case 'Character':
        p.setText(text.toUpperCase());
        p.setIndentFirstLine(180);
        p.setIndentStart(180);
        p.setIndentEnd(0);
        p.setSpacingBefore(12);
        p.setSpacingAfter(0);
        break;
      case 'Parenthetical':
        p.setIndentFirstLine(144);
        p.setIndentStart(144);
        p.setIndentEnd(144);
        p.setSpacingBefore(0);
        p.setSpacingAfter(0);
        break;
      case 'Dialogue':
        p.setIndentFirstLine(90);
        p.setIndentStart(90);
        p.setIndentEnd(90);
        p.setSpacingBefore(0);
        p.setSpacingAfter(12);
        break;
      case 'Transition':
        p.setText(text.toUpperCase());
        p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
        p.setIndentFirstLine(0);
        p.setIndentStart(0);
        p.setIndentEnd(0);
        p.setSpacingBefore(12);
        p.setSpacingAfter(12);
        break;
    }
  });

  for (let i = oldChildCount - 1; i >= 0; i--) {
    try {
      body.removeChild(body.getChild(i));
    } catch (e) {}
  }

  doc.saveAndClose();
  return 'Synced ' + validElements.length + ' lines';
}