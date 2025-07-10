const express = require("express");
const bodyParser = require("body-parser");
const pptxgenjs = require("pptxgenjs");
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

// Load background images from assets folder
const assetsDir = path.join(__dirname, "assets");
const bg1 = fs.readFileSync(path.join(assetsDir, "bg1.png"));
const bg2 = fs.readFileSync(path.join(assetsDir, "bg2.png"));

// Track usage stats
let pptxGeneratedToday = 0;

// Hardcoded styles for the Govplace template
const styles = {
  titleSlide: {
    title: {
      x: 0.5,            // inches from left
      y: 1.5,            // inches from top
      w: 9,              // width in inches
      h: 1.5,            // height in inches
      fontSize: 72,      // large title font
      color: "FFFFFF",   // white text
      bold: true,
      fontFace: "Arial", // sans-serif font
      align: "center"
    },
    subtitle: {
      x: 0.5,
      y: 3.5,
      w: 9,
      h: 1,
      fontSize: 36,      // medium subtitle font
      color: "FFFFFF",   // white text
      fontFace: "Arial",
      align: "center"
    }
  },
  contentSlide: {
    title: {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1,
      fontSize: 44,      // prominent content title
      color: "17375E",   // dark blue for contrast
      bold: true,
      fontFace: "Arial",
      align: "left"
    },
    content: {
      x: 0.5,
      y: 1.5,
      w: 9,
      h: 4.5,
      fontSize: 24,      // readable content font
      color: "333333",   // dark gray for text
      fontFace: "Arial",
      bullet: true,      // bulleted list
      wrap: true         // wrap long text
    }
  },
  footer: {
    x: 0,
    y: 6.7,
    w: "100%",
    fontSize: 14,        // small footer font
    color: "888888",     // light gray for subtlety
    align: "center",
    fontFace: "Arial"
  }
};

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
    const content = lines.join("\n").replace(/^- /gm, "• ");
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
        <link rel="stylesheet" type="text/css" href="/main.css">
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

    const pptx = new pptxgenjs();
    const slides = parseSlides(script);

    slides.forEach((slide, idx) => {
      const pptSlide = pptx.addSlide();
      // Apply background based on slide type
      pptSlide.background = { data: idx === 0 ? bg1 : bg2 };

      if (idx === 0) {
        // Title slide
        pptSlide.addText(slide.title, styles.titleSlide.title);
        pptSlide.addText(slide.content, styles.titleSlide.subtitle);
      } else {
        // Content slide
        pptSlide.addText(slide.title, styles.contentSlide.title);
        if (slide.content) {
          pptSlide.addText(slide.content, styles.contentSlide.content);
        }
      }
      // Add footer to every slide
      pptSlide.addText("Govplace Confidential", styles.footer);
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
