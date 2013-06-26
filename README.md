# MsOfficeWordHTMLWriter
  
  Create doc files in MHTML format

## Usage

    var MsOfficeWordHTMLWriter = require('./lib/MsOfficeWordHTMLWriter.js');

    var doc = new MsOfficeWordHTMLWriter({
      title : "My new doc",
      // WordDocument : { View:'Print' }
    });

    doc.write(
      "<p>hello, world</p>", 
      doc.page_break(),
      "<p>hello from another page</p>"
    );

    doc.create_section( {
      page: {
        size: "21.0cm 29.7cm",
        margin: "1.2cm 2.4cm 2.3cm 2.4cm"
      },
      header: "Section 2, page " + doc.field('PAGE') + " of " + doc.field('NUMPAGES'), 
      footer: "printed at " + doc.field('PRINTDATE'),
      new_page: 1 // or 'left', or 'right'
    });

    doc.write("this is the second section, look at header/footer");

    var png_path = 'demo.png';

    doc.attach("demo.png", png_path, function(doc) {
      doc.write("<img src='files/demo.png'>");
      doc.write("<table><tr><td>cell1</td><td>cell2</td></tr></table>");
      doc.save_as("demo");
    });
