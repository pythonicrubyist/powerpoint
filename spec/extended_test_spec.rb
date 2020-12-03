require 'powerpoint'

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do 
    @deck = Powerpoint::Presentation.new
    @deck.add_extended_intro 'Bicycle Of the Mind', 'created by Steve Jobs'
    @deck.save 'samples/pptx/sample.pptx' # Examine the PPTX file
  end

  it 'Create a PPTX file successfully.' do
    #@deck.should_not be_nil
  end
end