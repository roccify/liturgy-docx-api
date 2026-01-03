const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType, BorderStyle, HeightRule } = require('docx');

const cellBorders = {
  top: { style: BorderStyle.NONE, size: 0 },
  bottom: { style: BorderStyle.NONE, size: 0 },
  left: { style: BorderStyle.NONE, size: 0 },
  right: { style: BorderStyle.NONE, size: 0 }
};

function createParagraphsFromText(text) {
  // Split by double newline (paragraphs) and single newline (lines within paragraph)
  const paragraphs = [];
  const blocks = text.split('\n\n');
  
  blocks.forEach(block => {
    const lines = block.split('\n').filter(l => l.trim());
    if (lines.length > 0) {
      const runs = [];
      lines.forEach((line, idx) => {
        runs.push(new TextRun({ text: line, size: 24 }));
        if (idx < lines.length - 1) {
          runs.push(new TextRun({ break: 1 }));
        }
      });
      paragraphs.push(new Paragraph({
        spacing: { after: 100 },
        children: runs
      }));
    }
  });
  
  return paragraphs.length > 0 ? paragraphs : [new Paragraph({ children: [new TextRun({ text: text, size: 24 })] })];
}

function createLiturgyDocument(data) {
  const { subtitle, sections } = data;
  
  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      children: [new TextRun({ text: subtitle, size: 24 })]
    })
  ];

  sections.forEach(section => {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 300, after: 200 },
        children: [new TextRun({ text: section.name, bold: true, size: 28 })]
      })
    );

    const rows = [];
    
    // Reference row if exists
    if (section.latin.reference || section.slovenian.reference) {
      rows.push(
        new TableRow({
          height: { value: 300, rule: HeightRule.ATLEAST },
          children: [
            new TableCell({
              borders: cellBorders,
              width: { size: 50, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  spacing: { after: 200 },
                  children: [
                    new TextRun({
                      text: section.latin.reference || '',
                      italics: true,
                      size: 24
                    })
                  ]
                })
              ]
            }),
            new TableCell({
              borders: cellBorders,
              width: { size: 50, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  spacing: { after: 200 },
                  children: [
                    new TextRun({
                      text: section.slovenian.reference || '',
                      italics: true,
                      size: 24
                    })
                  ]
                })
              ]
            })
          ]
        })
      );
    }

    // Text row
    rows.push(
      new TableRow({
        children: [
          new TableCell({
            borders: cellBorders,
            width: { size: 50, type: WidthType.PERCENTAGE },
            verticalAlign: AlignmentType.TOP,
            children: createParagraphsFromText(section.latin.text || '')
          }),
          new TableCell({
            borders: cellBorders,
            width: { size: 50, type: WidthType.PERCENTAGE },
            verticalAlign: AlignmentType.TOP,
            children: createParagraphsFromText(section.slovenian.text || '')
          })
        ]
      })
    );

    children.push(
      new Table({
        columnWidths: [4680, 4680],
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: rows
      })
    );
  });

  return new Document({
    styles: {
      default: {
        document: {
          run: { font: "Times New Roman", size: 24 }
        }
      }
    },
    sections: [{
      properties: {
        page: {
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: children
    }]
  });
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed. Use POST.' });
  }

  try {
    const data = req.body;
    
    if (!data || !data.subtitle || !data.sections) {
      return res.status(400).json({ 
        error: 'Invalid request',
        required: ['subtitle', 'filename', 'sections']
      });
    }

    const doc = createLiturgyDocument(data);
    const buffer = await Packer.toBuffer(doc);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${data.filename || 'liturgy.docx'}"`);
    res.send(buffer);

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ 
      error: 'Failed to generate document',
      message: error.message
    });
  }
};
