var testCase = require('nodeunit').testCase,
    MsOfficeWordHTMLWriter = require("../lib/MsOfficeWordHTMLWriter.js");

exports["Base usage"] = {
    "Create document": function(test) {
        var doc = new MsOfficeWordHTMLWriter({
            title: "Demo",
            WordDocument: { 
                View: 'Print',
                Compatibility: {
                    DoNotExpandShiftReturn: ""
                }
            }
        });

        doc.write("hello, world");
        var br = doc.page_break('right');
        doc.write( br + 'new page after manual break' );
        doc.create_section( { new_page: 'right' } );
        doc.write('new page after section break');

        var content = doc.content();

        test.ok( /<w:View>Print<\/w:View>/g.test(content), "View => print" );
        test.ok( /<w:Compatibility><w:DoNotExpandShiftReturn \/>/g.test(content), "Compatibility => DoNotExpandShiftReturn" );

        var break_right_count = content.match(/page-break-before:right/g).length;

        test.equal( break_right_count, 2, "page break:right" );

        test.done();
    }
};