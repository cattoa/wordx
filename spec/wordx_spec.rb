require 'spec_helper'
require 'wordx'

describe '.Wordx' do

  it 'version number is 0.0.1' do
    expect(Wordx::VERSION).to eq '0.0.1'
  end

  before :each do
    @documents = Wordx::Documents.new()
    @styles = @documents.list_styles()
  end

  it 'takes no parameters and returns a Wordx::Documents object' do
      expect(@documents).to be_an_instance_of Wordx::Documents
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

end
