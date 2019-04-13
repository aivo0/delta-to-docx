const { Document, Paragraph, Packer } = require("docx");
var express = require("express");
const testContent = require("./testContent");

var app = express();

app.get("/", function(req, res) {
  const doc = new Document();

  const paragraph = new Paragraph("Hello World");

  doc.addParagraph(paragraph);
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
