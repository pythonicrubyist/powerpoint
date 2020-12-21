require 'powerpoint'

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do 
    @deck = Powerpoint::Presentation.new
    @deck.add_ranking_slide 'Bicycle Of the Mind','samples/images/sample_png.png', nil
    @deck.save 'samples/pptx/sample.pptx' # Examine the PPTX file
  end

  it 'Create a PPTX file successfully.' do
    #@deck.should_not be_nil
  end
end