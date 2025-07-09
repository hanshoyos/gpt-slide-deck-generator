const express = require('express');
const bodyParser = require('body-parser');
const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');

// Setup app and middleware
const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));

// Path to assets (logos and backgrounds)
const logoPath = './assets/GP-Logo-Facebook-04-01.png';
const bg1Path = './assets/bg1.png'; // Background for title slide
const bg2Path = './assets/bg2.png'; // Background for content slide

// Track the number of PowerPoints generated today
let pptxGeneratedToday = 0;
const currentDate = new Date().toLocaleDateString();

// Serve the home page with the form
app.get('/', (req, res) => {const express = require('express');
const bodyParser = require('body-parser');
const PptxGenJS = require('pptxgenjs');
const path = require('path');
const moment = require('moment-timezone');  // Import moment-timezone

// Setup app and middleware
const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));

// Path to assets (logos and backgrounds)
const logoPath = './assets/bbq.png'; // Logo image (assuming you have BBQ logo in assets folder)
const bg1Path = './assets/bg1.png'; // Background for title slide
const bg2Path = './assets/bg2.png'; // Background for content slide

// Track the number of PowerPoints generated today
let pptxGeneratedToday = 0;
const currentDate = moment().tz('America/New_York').format('YYYY-MM-DD'); // Get the current date in ET

// Serve the home page with the form
app.get('/', (req, res) => {
  const currentTimeET = moment().tz('America/New_York').format('YYYY-MM-DD HH:mm:ss'); // Get current time in ET

  res.send(`
    <html>
      <head>
        <title>Govplace PPTX Generator</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 40px; }
          textarea { width: 95%; }
          button { padding: 8px 20px; font-size: 18px; }
        </style>
      </head>
      <body>
        <h2>Govplace PowerPoint Generator</h2>
        <form method="POST" action="/download">
          <textarea name="script" rows="12" cols="80" placeholder="Enter your script here (e.g., Slide 1: Title)"></textarea><br/><br/>
          <button type="submit">Generate PowerPoint</button>
        </form>

        <hr>

        <footer>
          <p>Created by Hans Hoyos</p>
          <p>Current Time (ET): ${currentTimeET}</p>
          <p>PowerPoints generated today: ${pptxGeneratedToday}</p>
        </footer>
      </body>
    </html>
  `);
});

// Handle form submission and generate the PowerPoint
app.post('/download', (req, res) => {
  const { script } = req.body;

  if (!script || script.trim() === '') {
    return res.status(400).send('No script provided');
  }

  const pptx = new PptxGenJS();

  // Generate title slide
  const titleSlide = pptx.addSlide();
  titleSlide.background = { path: bg1Path };  // Set background for Title Slide
  titleSlide.addText('Govplace Solution Overview', {
    x: 0.5,
    y: 0.35,
    fontSize: 40,
    bold: true,
    color: '17375e',
    fontFace: 'Arial',
    align: 'center'
  });
  titleSlide.addImage({
    path: logoPath,
    x: 8.5,
    y: 0.2,
    w: 1.4,
    h: 0.7
  });

  // Process the input script directly
  const slidesContent = script.split('\n\n'); // Assuming each slide is separated by two newlines
  slidesContent.forEach((slideContent, idx) => {
    const slide = pptx.addSlide();
    slide.background = { path: idx === 0 ? bg1Path : bg2Path };  // Use bg1 for title slide, bg2 for others
    const [title, ...content] = slideContent.split('\n');

    slide.addText(title, {
      x: 0.5,
      y: 0.35,
      fontSize: 30,
      bold: true,
      color: '17375e',
      fontFace: 'Arial',
      align: 'left'
    });

    // Add content (if any)
    if (content.length > 0) {
      slide.addText(content.join('\n'), {
        x: 0.5,
        y: 1.3,
        fontSize: 18,
        color: '333333',
        fontFace: 'Arial',
        align: 'left'
      });
    }

    // Add footer text to each slide
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

  // Increment the PowerPoint generation counter
  pptxGeneratedToday++;

  // Generate the PowerPoint file
  const fileName = `Govplace_SlideDeck_${Date.now()}.pptx`;
  pptx.writeFile({ fileName }).then(() => {
    res.download(path.join(__dirname, fileName));
  }).catch((err) => {
    console.error('Error generating PowerPoint:', err);
    res.status(500).send("Error generating PowerPoint");
  });
});

// Start the server
app.listen(3000, () => {
  console.log('Server running on port 3000');
});

  res.send(`
    <html>
      <head>
        <title>Govplace PPTX Generator</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 40px; }
          textarea { width: 95%; }
          button { padding: 8px 20px; font-size: 18px; }
        </style>
      </head>
      <body>
        <h2>Govplace PowerPoint Generator</h2>
        <form method="POST" action="/download">
          <textarea name="script" rows="12" cols="80" placeholder="Enter your script here (e.g., Slide 1: Title)"></textarea><br/><br/>
          <button type="submit">Generate PowerPoint</button>
        </form>

        <hr>

        <footer>
          <p>Created by Hans Hoyos</p>
          <p>Date: ${new Date().toLocaleString()}</p>
          <p>PowerPoints generated today: ${pptxGeneratedToday}</p>
        </footer>
      </body>
    </html>
  `);
});

// Handle form submission and generate the PowerPoint
app.post('/download', (req, res) => {
  const { script } = req.body;

  if (!script || script.trim() === '') {
    return res.status(400).send('No script provided');
  }

  const pptx = new PptxGenJS();

  // Generate title slide
  const titleSlide = pptx.addSlide();
  titleSlide.background = { path: bg1Path };  // Set background for Title Slide
  titleSlide.addText('Govplace Solution Overview', {
    x: 0.5,
    y: 0.35,
    fontSize: 40,
    bold: true,
    color: '17375e',
    fontFace: 'Arial',
    align: 'center'
  });
  titleSlide.addImage({
    path: logoPath,
    x: 8.5,
    y: 0.2,
    w: 1.4,
    h: 0.7
  });

  // Process the input script directly
  const slidesContent = script.split('\n\n'); // Assuming each slide is separated by two newlines
  slidesContent.forEach((slideContent, idx) => {
    const slide = pptx.addSlide();
    slide.background = { path: idx === 0 ? bg1Path : bg2Path };  // Use bg1 for title slide, bg2 for others
    const [title, ...content] = slideContent.split('\n');

    slide.addText(title, {
      x: 0.5,
      y: 0.35,
      fontSize: 30,
      bold: true,
      color: '17375e',
      fontFace: 'Arial',
      align: 'left'
    });

    // Add content (if any)
    if (content.length > 0) {
      slide.addText(content.join('\n'), {
        x: 0.5,
        y: 1.3,
        fontSize: 18,
        color: '333333',
        fontFace: 'Arial',
        align: 'left'
      });
    }

    // Add logo to each slide
    slide.addImage({
      path: logoPath,
      x: 8.5,
      y: 0.2,
      w: 1.4,
      h: 0.7
    });

    // Add footer text to each slide
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

  // Increment the PowerPoint generation counter
  pptxGeneratedToday++;

  // Generate the PowerPoint file
  const fileName = `Govplace_SlideDeck_${Date.now()}.pptx`;
  pptx.writeFile({ fileName }).then(() => {
    res.download(path.join(__dirname, fileName));
  }).catch((err) => {
    console.error('Error generating PowerPoint:', err);
    res.status(500).send("Error generating PowerPoint");
  });
});

// Start the server
app.listen(3000, () => {
  console.log('Server running on port 3000');
});
