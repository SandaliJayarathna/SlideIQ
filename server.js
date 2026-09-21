const express = require("express");
const path = require("path");
const pptxgen = require("pptxgenjs");

const app = express();
const PORT = 3000;

app.use(express.json({ limit: "50mb" }));

app.use(express.static(__dirname));

app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "index.html"));
});

app.post("/generate-pptx", async (req, res) => {
    try {
        const data = req.body;

        const pptx = new pptxgen();

        pptx.layout = "LAYOUT_WIDE";
        pptx.author = "SlideIQ";
        pptx.title = "Office Farewell Presentation";

        // Slide 1
        let slide = pptx.addSlide();

        slide.background = { color: "F7F1E8" };

        slide.addText("CELEBRATING THE JOURNEY OF", {
            x: 1,
            y: 1.2,
            w: 11,
            h: 0.4,
            fontSize: 16,
            bold: true,
            align: "center",
            color: "7A6652"
        });

        slide.addText(data.personName || "Person Name", {
            x: 1,
            y: 2,
            w: 11,
            h: 0.8,
            fontSize: 38,
            bold: true,
            align: "center",
            color: "4D3B2B"
        });

        slide.addText(data.jobPosition || "", {
            x: 1,
            y: 3,
            w: 11,
            h: 0.4,
            fontSize: 20,
            align: "center",
            color: "6D6259"
        });

        slide.addText(data.companyName || "", {
            x: 1,
            y: 3.55,
            w: 11,
            h: 0.4,
            fontSize: 18,
            align: "center",
            color: "6D6259"
        });

        // Slide 2
        slide = pptx.addSlide();

        slide.addText("A JOURNEY OF DEDICATION", {
            x: 1,
            y: 1,
            w: 11,
            h: 0.5,
            fontSize: 28,
            bold: true,
            align: "center"
        });

        slide.addText(String(data.yearsWorked || "0"), {
            x: 1,
            y: 2,
            w: 11,
            h: 1,
            fontSize: 52,
            bold: true,
            align: "center"
        });

        slide.addText("Years of valuable service", {
            x: 1,
            y: 3.3,
            w: 11,
            h: 0.5,
            fontSize: 20,
            align: "center"
        });

        // Slide 3
        slide = pptx.addSlide();

        slide.addText("Achievements", {
            x: 1,
            y: 1,
            w: 11,
            h: 0.6,
            fontSize: 30,
            bold: true,
            align: "center"
        });

        slide.addText(data.achievements || "No achievements added.", {
            x: 1.5,
            y: 2,
            w: 10,
            h: 3,
            fontSize: 20,
            align: "center",
            valign: "mid",
            margin: 0.2
        });

        // Slide 4
        slide = pptx.addSlide();

        slide.addText("Special Memories", {
            x: 1,
            y: 1,
            w: 11,
            h: 0.6,
            fontSize: 30,
            bold: true,
            align: "center"
        });

        slide.addText(
            data.memories || "Special memories will appear here.",
            {
                x: 1.5,
                y: 2,
                w: 10,
                h: 3,
                fontSize: 20,
                align: "center",
                valign: "mid",
                margin: 0.2
            }
        );

        // Slide 5
        slide = pptx.addSlide();

        slide.addText("Farewell Message", {
            x: 1,
            y: 1,
            w: 11,
            h: 0.6,
            fontSize: 30,
            bold: true,
            align: "center"
        });

        slide.addText(
            data.farewellMessage ||
            "Wishing you all the best for the future!",
            {
                x: 1.5,
                y: 2,
                w: 10,
                h: 3,
                fontSize: 22,
                align: "center",
                valign: "mid",
                margin: 0.2
            }
        );

        // Slide 6 - Photos
        if (data.photos && data.photos.length > 0) {

            slide = pptx.addSlide();

            slide.addText("Memorable Moments", {
                x: 1,
                y: 0.5,
                w: 11,
                h: 0.5,
                fontSize: 30,
                bold: true,
                align: "center"
            });

            const positions = [
                { x: 0.7, y: 1.3 },
                { x: 4.35, y: 1.3 },
                { x: 8.0, y: 1.3 },
                { x: 0.7, y: 4.0 },
                { x: 4.35, y: 4.0 },
                { x: 8.0, y: 4.0 }
            ];

            for (let i = 0; i < data.photos.length && i < 6; i++) {

                slide.addImage({
                    data: data.photos[i],
                    x: positions[i].x,
                    y: positions[i].y,
                    w: 3,
                    h: 2.2
                });

            }
        }

        const buffer = await pptx.write({
            outputType: "nodebuffer"
        });

        res.set({
            "Content-Type":
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "Content-Disposition":
                'attachment; filename="office-farewell-presentation.pptx"',
            "Content-Length": buffer.length
        });

        res.send(buffer);

    } catch (error) {

        console.error(error);

        res.status(500).json({
            message: "Could not generate PowerPoint."
        });

    }
});

app.listen(PORT, () => {
    console.log(`SlideIQ running at http://localhost:${PORT}`);
});