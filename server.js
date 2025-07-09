// server.js
const express = require('express');
const bodyParser = require('body-parser');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));

// Serve the home page with a simple web form
app.get('/', (req, res) => {
  res.send(`
    <html>
      <head>
        <title>PPTX Generator</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 40px; }
          textarea { width: 95%; }
          button { padding: 8px 20px; font-size: 18px; }
        </style>
      </head>
      <body>
        <h2>Paste your script below to generate a PowerPoint:</h2>
        <form method="POST" action="/download" enctype="application/x-www-form-urlencoded">
          <textarea name="script" rows="12" cols="80" placeholder="Slide 1: Welcome\nThis is the first slide.\n\nSlide 2: Agenda\nFirst point\nSecond point"></textarea><br/><br/>
          <button type="submit">Generate PowerPoint</button>
        </form>
        <p>After clicking, your browser will download a .pptx file you can open in PowerPoint.</p>
      </body>
    </html>
  `);
});

// Endpoint to accept script and return the generated .pptx file
app.post('/download', async (req, res) => {
  const { script } = req.body;
  if (!script || !script.trim()) {
    return res.send("No script provided! Please go back and enter your slide content.");
  }

  const pptx = new PptxGenJS();

  // Split on blank lines or "Slide X:" labels (case-insensitive)
  const slides = script.split(/\n\s*\n|Slide \d+:/i).filter(s => s.trim());
  slides.forEach((slideText, idx) => {
    const slide = pptx.addSlide();
    const [title, ...body] = slideText.trim().split('\n');
    slide.addText(title || `Slide ${idx + 1}`, { x: 0.5, y: 0.5, fontSize: 24, bold: true });
    if (body.length > 0) {
      slide.addText(body.join('\n'), { x: 0.5, y: 1.2, fontSize: 18 });
    }
  });

  const fileName = `SlideDeck_${Date.now()}.pptx`;
  const filePath = path.join(__dirname, fileName);

  try {
    await pptx.writeFile({ fileName }); // Saves file locally
    res.download(filePath, fileName, (err) => {
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath); // Clean up after sending
    });
  } catch (err) {
    console.error('Error generating PPTX:', err);
    res.status(500).send("An error occurred while generating your PowerPoint file.");
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`PPTX Generator running on port ${PORT}`));
