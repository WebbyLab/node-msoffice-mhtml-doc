 use MsOffice::Word::HTML::Writer;
  my $doc = MsOffice::Word::HTML::Writer->new(
    title        => "My new doc",
    WordDocument => {View => 'Print'},
  );
  
  $doc->write("<p>hello, world</p>", 
              $doc->page_break, 
              "<p>hello from another page</p>");
  
  $doc->create_section(
    page => {size   => "21.0cm 29.7cm",
             margin => "1.2cm 2.4cm 2.3cm 2.4cm"},
    header => sprintf("Section 2, page %s of %s", 
                                  $doc->field('PAGE'), 
                                  $doc->field('NUMPAGES')),
    footer => sprintf("printed at %s", 
                                  $doc->field('PRINTDATE')),
    new_page => 1, # or 'left', or 'right'
  );

  $doc->write("this is the second section, look at header/footer");
  
$png_path = 'demo.png';

  $doc->attach("demo.png", $png_path);
  $doc->write("<img src='files/demo.png'>");
  $doc->write("<table><tr><td>cell1</td><td>cell2</td></tr></table>");
  
  $doc->save_as("demo");