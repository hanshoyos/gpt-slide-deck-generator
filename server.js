const express = require('express');
const bodyParser = require('body-parser');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const LOGO_PATH = './GP-Logo-Facebook-04-01.png';  // Local file in your repo root

const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));

const serverStartTime = new Date();
let usageCount = {};
function getTodayKey() {
  const d = new Date();
  return d.toISOString().slice(0, 10);
}

app.get('/', (req, res) => {
  const todayKey = getTodayKey();
  const countToday = usageCount[todayKey] || 0;

  const lastUpdated = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/New_York',
    year: 'numeric',
    month: 'long',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true
  }).format(serverStartTime);

  res.send(`
    <html>
      <head>
        <title>Govplace PPTX Generator</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 40px; }
          textarea { width: 95%; }
          button { padding: 8px 20px; font-size: 18px; }
          .info { color: gray; margin-bottom: 15px; }
        </style>
      </head>
      <body>
        <h2>Govplace Slide Deck Generator</h2>
        <div class="info">
          <p><strong>Server last started:</strong> ${lastUpdated} (ET)</p>
          <p><strong>Number of decks generated today:</strong> ${countToday}</p>
        </div>
        <form method="POST" action="/download" enctype="application/x-www-form-urlencoded">
          <textarea name="script" rows="12" cols="80" placeholder="Slide 1: Title\nIntro line\n\nSlide 2: Headline\n- Bullet 1\n- Bullet 2"></textarea><br/><br/>
          <button type="submit">Generate PowerPoint</button>
        </form>
        <p style="color:gray;font-size:14px;">Slides will match Govplace branding and include your logo.</p>
      </body>
    </html>
  `);
});

app.post('/download', async (req, res) => {
  const { script } = req.body;
  if (!script || !script.trim()) {
    return res.send("No script provided! Please go back and enter your slide content.");
  }

  const todayKey = getTodayKey();
  usageCount[todayKey] = (usageCount[todayKey] || 0) + 1;

  const pptx = new PptxGenJS();
  const slides = script.split(/\n\s*\n|Slide \d+:/i).filter(s => s.trim());
  slides.forEach((slideText, idx) => {
    const slide = pptx.addSlide();
    slide.background = { fill: 'FFFFFF' };

    const [title, ...body] = slideText.trim().split('\n');
    slide.addText(title || `Slide ${idx + 1}`, {
      x: 0.5,
      y: 0.35,
      fontSize: 30,
      bold: true,
      color: '17375e',
      fontFace: 'Arial'
    });

    slide.addShape(PptxGenJS.ShapeType.rect, {
      x: 0.5,
      y: 1.1,
      w: 4.5,
      h: 0.07,
      fill: { color: '25d1db' },
      line: 'none'
    });

    if (body.length > 0) {
      slide.addText(body.join('\n'), {
        x: 0.5,
        y: 1.3,
        w: 8.5,
        h: 5,
        fontSize: 18,
        color: '2E2E2E',
        fontFace: 'Arial'
      });
    }

    slide.addImage({ path: LOGO_PATH, x: 8.5, y: 0.2, w: 1.4, h: 0.7 });

    slide.addText('Govplace Confidential', {
      x: 0,
      y: 6.7,
      w: '100%',
      fontSize: 14,
      color: '888888',
      align: 'center',
      fontFace: 'Arial'
    });
  });

  const fileName = `Govplace_Deck_${Date.now()}.pptx`;
  const filePath = path.join(__dirname, fileName);

  try {
    await pptx.writeFile({ fileName });
    res.download(filePath, fileName, (err) => {
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    });
  } catch (err) {
    console.error('Error generating PPTX:', err);
    res.status(500).send("An error occurred while generating your PowerPoint file.");
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Govplace PPTX Generator running on port ${PORT}`));
