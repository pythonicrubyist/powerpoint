require 'powerpoint'

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do
    # Powerpoint.decompress_pptx 'test2.pptx', 'test2'
    # Powerpoint.decompress_pptx 'test1.pptx', 'test1'
    # Powerpoint.compress_pptx 'extract', 'x13.pptx'
    @deck = Powerpoint::Presentation.new 
    @deck.add_intro 'Bicycle Of the Mind', 'created by Steve Jobs'
    @deck.add_textual_slide 'Why Mac?', ['Its cool!', 'Its light!']
    @deck.add_textual_slide 'Why Iphone?', ['Its fast!', 'Its cheap!']
    @deck.add_pictorial_slide 'Logo', 'image1.jpg'
    @deck.add_textual_slide 'Why Android?', ['Its great!', 'Its sweet!']
    @deck.save 'sample.pptx'
  end

  it 'Create a PPTX file successfully.' do
    #@deck.should_not be_nil
  end
end
