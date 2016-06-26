require 'spec_helper'
require 'wordx'

describe '.Wordx' do

  it 'version number is 0.0.1' do
    expect(Wordx::VERSION).to eq '0.0.1'
  end

  before :each do
    @document = Wordx::Document.new()
    @styles = @document.list_styles()
  end

  it 'takes no parameters and returns a Wordx::Documents object' do
      expect(@document).to be_an_instance_of Wordx::Document
  end

  it 'documents should contain 2 Wordx::Styles' do
    expect(@styles.count).to eq 2
  end

  it 'style should contain DefaultText' do
    expect(@styles).to include(:DefaultText)
  end

  it 'style should contain Normal' do
    expect(@styles).to include(:Normal)
  end

  it 'add paragraph to document' do
    @para_key = @document.new_paragraph()
    expect(@para_key).to eq :para_001
  end

  it 'add text to paragrah' do
    @document.add_text(@para_key, "Hello World")
    expect(@document.get_text(@para_key)).to eq "Hello World"
  end

end
