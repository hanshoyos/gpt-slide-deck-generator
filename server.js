const express = require('express');
const bodyParser = require('body-parser');
const PptxGenJS = require('pptxgenjs');
const path = require('path');

const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));

// Path to assets (logos and backgrounds)
const logoPath = './assets/GP-Logo-Facebook-04-01.png';
const bg1Path = './assets/bg1.png'; // Background for title slide
const bg2Path = './assets/bg2.png'; // Background for content slide

app.post('/download', (req, res) => {
  const pptx = new PptxGenJS();

  // Example for Title Slide
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

  // Example for Content Slide
  const contentSlide = pptx.addSlide();
  contentSlide.background = { path: bg2Path };  // Set background for Content Slide
  contentSlide.addText('Why Modernize Identity Now?', {
    x: 0.5,
    y: 0.35,
    fontSize: 30,
    bold: true,
    color: '17375e',
    fontFace: 'Arial',
    align: 'left'
  });
  contentSlide.addText('- Federal mandates (e.g., EO 14028, OMB M-22-09) require Zero Trust\n- Legacy identity systems slow down user access and increase risk\n- Cyber threats targeting identity are on the rise', {
    x: 0.5,
    y: 1.3,
    fontSize: 18,
    color: '333333',
    fontFace: 'Arial',
    align: 'left'
  });

  // Generate the file
  const fileName = `Govplace_SlideDeck_${Date.now()}.pptx`;
  pptx.writeFile({ fileName }).then(() => {
    res.download(path.join(__dirname, fileName));
  }).catch((err) => {
    res.status(500).send("Error generating PowerPoint");
  });
});

app.listen(3000, () => {
  console.log('Server running on port 3000');
});
