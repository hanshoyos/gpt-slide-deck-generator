const express = require("express");
const bodyParser = require("body-parser");
const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const moment = require("moment-timezone");

const app = express();
const port = process.env.PORT || 3000;

app.use(bodyParser.json({ limit: "5mb" }));
app.use(express.static(path.join(__dirname, "public")));

// Ensure outputs directory exists
const outputsDir = path.join(__dirname, "outputs");
if (!fs.existsSync(outputsDir)) fs.mkdirSync(outputsDir, { recursive: true });

// Load background images
const assetsDir = path.join(__dirname, "assets");
const bg1Path = path.join(assetsDir, "bg1.png");
const bg2Path = path.join(assetsDir, "bg2.png");

// Verify background images exist and are readable
if (!fs.existsSync(bg1Path) || !fs.existsSync(bg2Path)) {
  console.error("Background images (bg1.png or bg2.png) are missing in the assets folder.");
  process.exit(1);
}

// Usage tracking
let pptxGeneratedToday = 0;

// Style presets with base values (fontSize will be set dynamically)
const baseStyles = {
  titleSlide: {
    title: {
      x: 0.97, y: 4.06, w: 8.32, h: 0.66,
      color: "005670", bold: true,
      fontFace: "Arial", align: "center", valign: "middle",
      margin: [0.05, 0.1, 0.05, 0.1], wrap: true, fit: true
    },
    subtitle: {
      x: 0.63, y: 4.72, w: 9, h: 0.77,
      color: "A9A9A9", fontFace: "Arial",
      align: "center", valign: "middle",
      margin: [0.05, 0.1, 0.05, 0.1], wrap: true, fit: true,
      fontSize: 32
    }
  },
  contentSlide: {
    title: {
      x: 0.7, y: 0.39, w: 3.5, h: 2.17,
      color: "005670", bold: true, fontFace: "Arial",
      align: "left", valign: "middle",
      margin: [0.05, 0.1, 0.05, 0.1], wrap: true, fit: true
    },
    content: {
      x: 4.59, y: 0.39, w: 5.03, h: 4.73,
      fontSize: 24, color: "333333", fontFace: "Arial",
      bullet: { style: "circle", indent: 0.25 },
      align: "left", valign: "middle",
      margin: [0.05, 0.1, 0.05, 0.1], wrap: true, fit: true
    }
  },
  footer: {
    x: 0, y: 5.48, w: 10, h: 0,
    fontSize: 12, color: "A9A9A9", fontFace: "Arial",
    align: "center", valign: "middle",
    margin: [0.05, 0.1, 0.05, 0.1], wrap: true, fit: true
  }
};

// Dynamic font size helper
function getDynamicFontSize(text, max, mid, min, maxLen, midLen) {
  if (!text) return max;
  if (text.length <= maxLen) return max;
  if (text.length <= midLen) return mid;
  return min;
}

// Get current time in Eastern Time
function getCurrentTimeET() {
  return moment().tz("America/New_York").format("MM/DD/YYYY hh:mm:ss A");
}

// Parse script into slides
function parseSlides(scriptText) {
  const slideRegex = /Slide \d+:\s*(.+?)(?=(Slide \d+:|$))/gs;
  const matches = [...scriptText.matchAll(slideRegex)];
  return matches.map((match) => {
    const slideText = match[1].trim();
    const lines = slideText.split(/\r?\n/).filter(Boolean);
    const title = lines.shift() || "Untitled";
    const content = lines.join("\n").replace(/^- /gm, "");
    return { title, content };
  });
}

// Serve the web form
app.get("/", (req, res) => {
  const currentTimeET = getCurrentTimeET();
  res.send(`
    <html>
      <head>
        <title>Govplace PPTX Generator</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; }
          h2 { color: #005670; }
          textarea { width: 100%; max-width: 600px; }
          button { background: #005670; color: white; padding: 10px; border: none; cursor: pointer; }
          button:hover { background: #003d50; }
          footer { margin-top: 20px; color: #A9A9A9; font-size: 12px; }
        </style>
      </head>
      <body>
        <h2>Govplace PowerPoint Generator</h2>
        <form id="generateForm" method="POST" action="/createSlideDeck">
          <textarea name="script" rows="12" cols="80" placeholder="Paste your script here (e.g., Slide 1: Title)"></textarea><br/><br/>
          <button type="submit">Generate PowerPoint</button>
        </form>
        <div id="status"></div>
        <hr>
        <footer>
          <p>Created by Hans Hoyos</p>
          <p>Current Time (ET): ${currentTimeET}</p>
          <p>PowerPoints generated today: ${pptxGeneratedToday}</p>
        </footer>
        <script>
          document.getElementById("generateForm").addEventListener("submit", async (e) => {
            e.preventDefault();
            const formData = new FormData(e.target);
            const script = formData.get("script");
            const statusDiv = document.getElementById("status");
            statusDiv.innerHTML = "Generating PowerPoint...";
            try {
              const response = await fetch("/createSlideDeck", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ script }),
              });
              const data = await response.json();
              if (data.fileUrl) {
                statusDiv.innerHTML = "PowerPoint generated! Downloading...";
                window.location.href = data.fileUrl;
              } else {
                statusDiv.innerHTML = "Error: " + data.error;
              }
            } catch (err) {
              statusDiv.innerHTML = "Error: " + err.message;
            }
          });
        </script>
      </body>
    </html>
  `);
});

// Handle PPTX generation
app.post("/createSlideDeck", async (req, res) => {
  try {
    const { script } = req.body;
    if (!script) return res.status(400).json({ error: "Script text is required." });

    const pptx = new PptxGenJS();
    const slides = parseSlides(script);

    slides.forEach((slide, idx) => {
      const pptSlide = pptx.addSlide();
      // Apply background based on slide type
      pptSlide.background = { path: idx === 0 ? bg1Path : bg2Path };

      if (idx === 0) {
        // Title slide
        const titleFontSize = getDynamicFontSize(slide.title, 60, 48, 36, 36, 54);
        const titleStyle = { ...baseStyles.titleSlide.title, fontSize: titleFontSize };
        pptSlide.addText(slide.title, titleStyle);

        if (slide.content) {
          // Optional: Adjust subtitle font size dynamically as well if desired
          pptSlide.addText(slide.content, baseStyles.titleSlide.subtitle);
        }
      } else {
        // Content slide
        const titleFontSize = getDynamicFontSize(slide.title, 40, 32, 24, 40, 70);
        const contentTitleStyle = { ...baseStyles.contentSlide.title, fontSize: titleFontSize };
        pptSlide.addText(slide.title, contentTitleStyle);

        if (slide.content) {
          pptSlide.addText(slide.content, baseStyles.contentSlide.content);
        }
      }
      // Add footer
      pptSlide.addText("Govplace Confidential", baseStyles.footer);
    });

    pptxGeneratedToday++;
    const filename = `Govplace_Slide_Deck_${Date.now()}.pptx`;
    const filepath = path.join(outputsDir, filename);
    await pptx.writeFile({ fileName: filepath });

    const fileUrl = `${req.protocol}://${req.get("host")}/outputs/${filename}`;
    res.json({ fileUrl });
  } catch (error) {
    console.error("Error generating PowerPoint:", error);
    res.status(500).json({ error: "Failed to generate PowerPoint." });
  }
});

// Serve generated files
app.use("/outputs", express.static(outputsDir));

app.listen(port, () => {
  console.log(`Govplace PPTX Generator running on port ${port} at ${getCurrentTimeET()}`);
});
