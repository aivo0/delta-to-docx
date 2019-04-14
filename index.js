const {
  Document,
  Paragraph,
  Packer,
  TextRun,
  Numbering,
  Indent,
  Hyperlink
} = require("docx");
const express = require("express");
const fs = require("fs");
const testContent = require("./testContent");

var app = express();

app.get("/", function(req, res) {
  const author = req.query.author ? req.query.author : "";
  const title = req.query.title ? req.query.title : "";
  const doc = new Document({ creator: author, title });
  // To support numbered paragraphs
  const numbering = new Numbering();
  const abstractNum = numbering.createAbstractNumbering();
  abstractNum
    .createLevel(0, "upperRoman", "%1", "start")
    .addParagraphProperty(new Indent(720, 260));
  const concrete = numbering.createConcreteNumbering(abstractNum);
  // A temporary variable to hold onto multiple words before inserting
  let prg = new Paragraph();
  testContent.ops.forEach(op => {
    if (op.insert && typeof op.insert === "object" && op.insert.speaker) {
      doc.addParagraph(new Paragraph(""));
      doc
        .createParagraph()
        .createTextRun(op.insert.speaker + ":")
        .bold();
    } else if (
      op.insert &&
      typeof op.insert == "string" &&
      op.insert.includes("\n")
    ) {
      if (op.attributes && op.attributes.align) {
        switch (op.attributes.align) {
          case "right":
            prg.right();
            break;
          case "center":
            prg.center();
            break;
          case "justify":
            prg.justified();
            break;
        }
      }
      if (op.attributes && op.attributes.list) {
        if (op.attributes.list == "bullet") {
          prg.bullet();
        } else {
          prg.setNumbering(concrete, 0);
        }
      } else if (op.attributes && op.attributes.header) {
        switch (op.attributes.header) {
          case 1:
            prg.heading1();
            break;
          case 2:
            prg.heading2();
            break;
          case 3:
            prg.heading3();
            break;
        }
      }
      doc.addParagraph(prg);
      prg = new Paragraph();
    } else if (op.insert) {
      if (op.insert.image) {
        // The string has a pattern "data:image/png;base64,...
        fs.writeFileSync(
          "./temp-image.jpg",
          op.insert.image.replace(/^data:image\/\w+;base64,/, ""),
          {
            encoding: "base64"
          }
        );
        doc.createImage(fs.readFileSync("./temp-image.jpg"));
      } else if (op.attributes && op.attributes.link) {
        prg.addHyperLink(new Hyperlink(op.insert, 1, op.attributes.link));
      } else {
        const text = new TextRun(op.insert);
        if (op.attributes) {
          if (op.attributes.bold) text.bold();
          if (op.attributes.italic) text.italic();
          if (op.attributes.strike) text.strike();
          if (op.attributes.underline) text.underline();
          if (op.attributes.color) text.color(op.attributes.color);
          /* if (op.attributes.font) text.font(op.attributes.font); */
        }
        prg.addRun(text);
      }
    }
  });

  const packer = new Packer(doc, undefined, undefined, numbering);

  packer.toBase64String(doc).then(b64string => {
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=My Document.docx"
    );
    res.send(Buffer.from(b64string, "base64"));
  });
});

app.listen(3333);
console.log("Listening on: localhost:3333");
