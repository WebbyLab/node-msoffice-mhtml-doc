var MsOfficeWordHTMLWriter = require( __dirname + '/../lib/MsOfficeWordHTMLWriter' );

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
doc.write( doc.page_break('right') + 'new page after manual break' );
doc.create_section( { new_page: 'right' } );
doc.write('new page after section break');

var content = doc.content();
var break_right_count = content.match(/page-break-before:right/g).length;

test('Create document', function(assert) {
    ok( /<w:View>Print<\/w:View>/g.test(content), "View => print" );
    ok( /<w:Compatibility><w:DoNotExpandShiftReturn \/>/g.test(content), "Compatibility => DoNotExpandShiftReturn" );
    equal( break_right_count, 2, "page break:right" );
});
