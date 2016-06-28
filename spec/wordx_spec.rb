require 'spec_helper'
require 'wordx'

RSpec.describe '.Wordx' do

  it 'version number is 0.0.1' do
    expect(Wordx::VERSION).to eq '0.0.1'
  end

  before (:context) do
    @document = Wordx::Document.new()
    @styles = @document.list_styles()
    @para_key = @document.new_paragraph()

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
    expect(@para_key).to eq :para_001
  end

  it 'paragraphs should contain a paragraph key of para_001' do
    keys = @document.list_paragraphs
    expect(keys).to include(:para_001)
  end

  it 'add text to paragrah par_001' do
    @document.add_text("Hello World",@para_key)
    expect(@document.get_text(@para_key)).to eq "Hello World"
  end
  it 'add text to current paragrah' do
    @document.add_text(" glad to know you")
    expect(@document.get_text()).to eq "Hello World glad to know you"
  end

   it 'add second paragraph to document' do
     @para_key_2 = @document.new_paragraph()
     expect(@para_key_2).to eq :para_002
   end

   it 'add text to paragraph 2' do
     @document.add_text("Third rock from the sun",@para_key_2)
     expect(@document.get_text(@para_key_2)).to eq "Third rock from the sun"
   end
   it 'add text to current paragrah' do
     @document.add_text(" is earth")
     expect(@document.get_text()).to eq "Third rock from the sun is earth"
   end
   it 'the first paragraph is till the same' do
     expect(@document.get_text(@para_key)).to eq "Hello World glad to know you"
   end
   it 'the default bold setting for the paragraph is false' do
     expect(@document.get_bold(@para_key)).to eq false
   end
   it 'set paragraph bold=true' do
     @document.set_bold(true,@para_key)
     expect(@document.get_bold(@para_key)).to eq true
   end

   it 'set paragraph to bold=false' do
     @document.set_bold(false,@para_key)
     expect(@document.get_bold(@para_key)).to eq false
   end

   it 'set paragraph bold=true' do
     @document.set_bold(true,@para_key)
     expect(@document.get_bold(@para_key)).to eq true
   end
   
   it 'set create document' do
     @document.create_document()
   end

 end
