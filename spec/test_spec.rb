require 'powerpoint'

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do
    #Powerpoint.decompress_pptx'blank.pptx', 'blank'
    @deck = Powerpoint::Presentation.new 'sample.pptx'
    @deck.add_intro 'Bicycle Of the Mind', 'created by Steve Jobs'
    @deck.add_textual_slide 'Why Mac?', ['Its cool!', 'Its light.']
    @deck.save
  end

  it 'Create a PPTX file successfully.' do
    @deck.should_not be_nil
  end
end
