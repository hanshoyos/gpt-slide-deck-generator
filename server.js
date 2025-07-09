// server.js
const express = require('express');
const bodyParser = require('body-parser');
const PptxGenJS = require("pptxgenjs");
const fs = require('fs');
const path = require('path');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const app = express();
app.use(bodyParser.json());

app.post('/createSlideDeck', async (req, res) => {
  const { script, access_token } = req.body;
  if (!script || !access_token) {
    return res.status(400).json({ error: 'Missing script or access_token.' });
  }

  // 1. Parse and generate the PowerPoint deck
  const pptx = new PptxGenJS();

  // Split script into slides. Expecting format:
  // Slide 1: Title
  // Bullet 1
  // Bullet 2
  // (Blank lines or Slide X: to split)
  const slides = script.split(/(?:\n\s*\n|Slide \d+:)/i).filter(s => s.trim());
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
    // 2. Write .pptx file locally
    await pptx.writeFile({ fileName });

    // 3. Upload to OneDrive via Microsoft Graph API
    const graphClient = Client.init({
      authProvider: (done) => done(null, access_token)
    });

    const fileStream = fs.createReadStream(filePath);
    // Create folder 'GeneratedSlides' if it doesn't exist
    try {
      await graphClient
        .api('/me/drive/root:/GeneratedSlides')
        .get();
    } catch {
      await graphClient
        .api('/me/drive/root/children')
        .post({ name: "GeneratedSlides", folder: {}, "@microsoft.graph.conflictBehavior": "rename" });
    }

    // Upload the pptx to /GeneratedSlides/ in user's OneDrive
    const uploadRes = await graphClient
      .api('/me/drive/root:/GeneratedSlides/' + fileName + ':/content')
      .putStream(fileStream);

    // 4. Get a shareable link
    const linkRes = await graphClient
      .api(`/me/drive/items/${uploadRes.id}/createLink`)
      .post({ type: 'view' });

    // Clean up local file
    fs.unlinkSync(filePath);

    // 5. Return the URL to client (ChatGPT Action)
    return res.json({ fileUrl: linkRes.link.webUrl });

  } catch (err) {
    console.error('Error creating or uploading PPTX:', err);
    // Clean up local file if exists
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    return res.status(500).json({ error: 'Failed to create or upload PowerPoint file.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Slide deck generator running on port ${PORT}`));
