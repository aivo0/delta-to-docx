const { Document, Paragraph, Packer } = require("docx");
var express = require("express");
const testContent = require("./testContent");

var app = express();

app.get("/", function(req, res) {
  const author = req.query.author ? req.query.author : "";
  const title = req.query.title ? req.query.title : "";
  const doc = new Document({ creator: author, title });
  let paragraph = "";
  testContent.ops.forEach(op => {
    if (op.insert && op.insert.speaker) {
      doc.addParagraph(new Paragraph(op.insert.speaker + ":"));
    } else if (op.insert && op.insert.includes("\n")) {
      paragraph = paragraph + op.insert;
      doc.addParagraph(new Paragraph(paragraph));
      paragraph = "";
    } else if (op.insert) {
      paragraph = paragraph + op.insert;
    }
  });

  const packer = new Packer();

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
