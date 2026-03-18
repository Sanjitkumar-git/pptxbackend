const PptxGenJS = require("pptxgenjs");

exports.handler = async (event) => {
  try {
    if (event.httpMethod !== "POST") {
      return {
        statusCode: 405,
        body: "Method Not Allowed",
      };
    }

    const body = JSON.parse(event.body || "{}");
    const { title, slides } = body;

    if (!slides || !Array.isArray(slides) || slides.length === 0) {
      return {
        statusCode: 400,
        body: JSON.stringify({ error: "Slides data missing" }),
      };
    }

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "NoteCraft AI";
    pptx.company = "NoteCraft AI";
    pptx.subject = title || "Presentation";
    pptx.title = title || "Presentation";
    pptx.lang = "en-US";

    slides.forEach((slideData, index) => {
      const slide = pptx.addSlide();
      const isTitleSlide = slideData.isTitleSlide === true;

      slide.background = {
        color: isTitleSlide ? "2B7DF2" : "0F2B4A",
      };

      slide.addText(isTitleSlide ? "TITLE SLIDE" : `SLIDE ${index + 1}`, {
        x: 0.3,
        y: 0.2,
        w: 1.9,
        h: 0.35,
        fontSize: 12,
        bold: true,
        color: isTitleSlide ? "2B7DF2" : "FFFFFF",
        fill: { color: isTitleSlide ? "FFFFFF" : "2B7DF2" },
        margin: 0.08,
        align: "center",
        valign: "mid",
        radius: 0.12,
      });

      if (isTitleSlide && slideData.subtitle) {
        slide.addText(slideData.subtitle, {
          x: 0.5,
          y: 0.75,
          w: 6.5,
          h: 0.4,
          fontSize: 18,
          italic: true,
          color: "FFFFFF",
        });
      }

      slide.addText(slideData.title || `Slide ${index + 1}`, {
        x: 0.5,
        y: isTitleSlide ? 1.15 : 0.85,
        w: slideData.diagramBase64 ? 6.3 : 12,
        h: 0.6,
        fontSize: isTitleSlide ? 24 : 22,
        bold: true,
        color: isTitleSlide ? "FFFFFF" : "00E5FF",
      });

      let startY = isTitleSlide ? 2.0 : 1.7;
      const content = Array.isArray(slideData.content) ? slideData.content : [];

      content.forEach((point, i) => {
        slide.addText(
          [
            {
              text: "• ",
              options: {
                bold: true,
                color: isTitleSlide ? "FFFFFF" : "00E5FF",
              },
            },
            {
              text: String(point),
              options: {
                color: "FFFFFF",
              },
            },
          ],
          {
            x: 0.6,
            y: startY + i * 0.55,
            w: slideData.diagramBase64 ? 6.2 : 11.8,
            h: 0.45,
            fontSize: 14,
            margin: 0,
            breakLine: false,
            valign: "top",
          }
        );
      });

      if (slideData.diagramBase64) {
        slide.addText("Diagram", {
          x: 7.5,
          y: 1.5,
          w: 2,
          h: 0.3,
          fontSize: 14,
          bold: true,
          color: "00E5FF",
        });

        slide.addShape(pptx.ShapeType.rect, {
          x: 7.3,
          y: 1.8,
          w: 5.4,
          h: 3.6,
          line: { color: "00E5FF", pt: 1 },
          fill: { color: "102540", transparency: 40 },
          radius: 0.12,
        });

        slide.addImage({
          data: slideData.diagramBase64,
          x: 7.4,
          y: 1.9,
          w: 5.2,
          h: 3.4,
        });
      }
    });

    const buffer = await pptx.write({ outputType: "nodebuffer" });

    return {
      statusCode: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="${title || "presentation"}.pptx"`,
      },
      body: buffer.toString("base64"),
      isBase64Encoded: true,
    };
  } catch (e) {
    console.error("PPTX generation error:", e);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: e.toString() }),
    };
  }
};