const express = require("express");
const bodyParser = require("body-parser");
const pptxgenjs = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const app = express();
const port = process.env.PORT || 3000;

app.use(bodyParser.json({ limit: "5mb" }));
app.use(express.static(path.join(__dirname, "public")));

const bg1 = fs.readFileSync(path.join(__dirname, "assets/bg1.png"));
const bg2 = fs.readFileSync(path.join(__dirname, "assets/bg2.png"));

let pptxGeneratedToday = 0;

function parseSlides(scriptText) {
  const slideRegex = /Slide \d+:\s*(.+?)(?=Slide \d+:|$)/gs;
  const matches = [...scriptText.matchAll(slideRegex)];

  return matches.map((match) => {
    const slideText = match[1].trim();
    const lines = slideText.split(/\r?\n/).filter(Boolean);
    const title = lines.shift() || "Untitled";
    const content = lines.join("\n").replace(/^- /gm, "• ");
    return { title, content };
  });
}

function getCurrentTimeET() {
  const options = {
    timeZone: "America/New_York",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  };
  return new Intl.DateTimeFormat("en-US", options).format(new Date());
}

app.get("/", (req, res) => {
  const currentTimeET = getCurrentTimeET();
  res.send(`
    <html>
      <head>
        <title>Govplace PPTX Generator</title>
        <link rel="stylesheet" type="text/css" href="/main.css">
      </head>
      <body>
        <h2>Govplace PowerPoint Generator</h2>
        <form method="POST" action="/createSlideDeck">
          <textarea name="script" rows="12" cols="80" placeholder="Paste your script here (e.g., Slide 1: Title)"></textarea><br/><br/>
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

app.post("/createSlideDeck", async (req, res) => {
  try {
    const { script } = req.body;
    if (!script) return res.status(400).json({ error: "Script text is required." });

    const pptx = new pptxgenjs();
    const slides = parseSlides(script);

    if (slides.length === 0) {
      return res.status(400).json({ error: "No slides parsed from script." });
    }

    slides.forEach((slide, idx) => {
      const pptSlide = pptx.addSlide();
      pptSlide.background = { data: idx === 0 ? bg1 : bg2 };

      // Title box styling from CSS
      pptSlide.addText(slide.title, {
        x: 1,
        y: 0.4,
        w: 8,
        h: 1.3,
        fontSize: 32,
        bold: true,
        color: "17375e",
        fontFace: "Arial",
      });

      // Content box styling from CSS
      if (slide.content) {
        pptSlide.addText(slide.content, {
          x: 1,
          y: 1.8,
          w: 8,
          h: 5,
          fontSize: 18,
          lineSpacing: 24,
          color: "333333",
          fontFace: "Arial",
          bullet: true,
          wrap: true,
          margin: 10,
        });
      }

      // Footer text (small, centered, lighter color)
      pptSlide.addText("Govplace Confidential", {
        x: 0,
        y: 6.7,
        w: "100%",
        fontSize: 14,
        color: "888888",
        align: "center",
        fontFace: "Arial",
      });
    });

    pptxGeneratedToday++;

    const filename = `Govplace_Slide_Deck_${Date.now()}.pptx`;
    const filepath = path.join(__dirname, "outputs", filename);
    await pptx.writeFile({ fileName: filepath });

    const fileUrl = `${req.protocol}://${req.get("host")}/outputs/${filename}`;
    res.json({ fileUrl });
  } catch (error) {
    console.error("Error generating PowerPoint:", error);
    res.status(500).json({ error: "Failed to generate PowerPoint." });
  }
});

app.use("/outputs", express.static(path.join(__dirname, "outputs")));

app.listen(port, () => {
  console.log(`Govplace PPTX Generator running on port ${port}`);
});
