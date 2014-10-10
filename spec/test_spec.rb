require 'powerpoint'

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do 
    @deck = Powerpoint::Presentation.new 
    @deck.add_intro 'Bicycle Of the Mind', 'created by Steve Jobs'
    @deck.add_textual_slide 'Why Mac?', ['Its cool!', 'Its light!']
    @deck.add_textual_slide 'Why Iphone?', ['Its fast!', 'Its cheap!']
    @deck.add_pictorial_slide 'JPG Logo', 'samples/images/sample_jpg.jpg'
    @deck.add_pictorial_slide 'PNG Logo', 'samples/images/sample_png.png'
    @deck.add_pictorial_slide 'GIF Logo', 'samples/images/sample_gif.gif', {x: 124200, y: 3356451, cx: 2895600, cy: 1013460}
    @deck.add_pictorial_slide 'SVG Logo', 'samples/images/sample_svg.svg'
    @deck.add_textual_slide 'Why Android?', ['Its great!', 'Its sweet!']
    @deck.add_table_slide "Table test", [["foo","bar","baz"],["1","2","3"]]
    @deck.save 'samples/pptx/sample.pptx' # Examine the PPTX file
  end

  it 'Create a PPTX file successfully.' do
    #@deck.should_not be_nil
  end
end
