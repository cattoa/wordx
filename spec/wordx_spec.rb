require 'spec_helper'
require 'wordx'

RSpec.describe '.Wordx' do

  it 'version number is 0.0.1' do
    expect(Wordx::VERSION).to eq '0.0.1'
  end

  before (:context) do
    @document = Wordx::Document.new()
    @styles = @document.list_styles()
    @paragraph = @document.new_paragraph()
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
    expect(@paragraph.key).to eq :para_001
  end

  it 'paragraphs should contain a paragraph key of para_001' do
    keys = @document.list_paragraphs
    expect(keys).to include(:para_001)
  end

  it 'add text to paragrah par_001' do
    @paragraph.text = "Hello World"
    expect(@paragraph.text).to eq "Hello World"
  end

  it 'add text to current paragrah' do
    @paragraph.text += " glad to know you"
    expect(@paragraph.text).to eq "Hello World glad to know you"
  end

   it 'add second paragraph to document' do
     paragraph = @document.new_paragraph()
     expect(paragraph.key).to eq :para_002
   end

   it 'add text to new paragraph ' do
     paragraph = @document.new_paragraph()
     paragraph.text = "Third rock from the sun"
     expect(paragraph.text).to eq "Third rock from the sun"
   end
   it 'add text to paragraph para_001' do
     paragraph = @document.paragraph(:para_001)
     paragraph.text += ".Hope you have a great day!"
     expect(paragraph.text).to eq "Hello World glad to know you.Hope you have a great day!"
   end

   it 'the second paragraph is still empty' do
     paragraph = @document.paragraph(:para_002)
     expect(paragraph.text).to eq ""
   end


   it 'the default bold setting for the paragraph is false' do
     paragraph = @document.paragraph(:para_001)
     expect(paragraph.bold).to eq false
   end


   it 'set paragraph bold=true' do
     paragraph = @document.paragraph(:para_001)
     paragraph.bold = true
     expect(paragraph.bold).to eq true
   end

   it 'set paragraph to bold=false' do
     paragraph = @document.paragraph(:para_001)
     paragraph.bold = false
     expect(paragraph.bold).to eq false
   end

   it 'create a table in the document' do
     rows = 3
     columns = 3
     paragraph = @document.paragraph(:para_001)
     table = paragraph.new_table(rows,columns)
     expect(table.length).to eq 3
   end


   it 'add text to cell(0,0) and get text from cell(0,0)' do
     rows = 3
     columns = 3
     paragraph = @document.paragraph(:para_001)
     table = paragraph.new_table(rows,columns)
     cell = table.cell(0,0)
     cell.text = "Hello World"
     expect(cell.text).to eql "Hello World"
   end

   it 'add text to cell(0,0) and get text from cell(0,1)' do
     rows = 3
     columns = 3
     paragraph = @document.paragraph(:para_001)
     table = paragraph.new_table(rows,columns)
     cell = table.cell(0,0)
     cell.text = "Hello World"
     cell = table.cell(0,1)
     expect(cell.text).to eql ""

   end

   it 'add cell(0,0) and bold should be false' do
     rows = 3
     columns = 3
     paragraph = @document.paragraph(:para_001)
     table = paragraph.new_table(rows,columns)
     cell = table.cell(0,0)
     expect(cell.bold).to eql false
   end


   it 'add cell(0,0) and set bold to true' do
     rows = 3
     columns = 3
     paragraph = @document.paragraph(:para_001)
     table = paragraph.new_table(rows,columns)
     cell = table.cell(0,0)
     cell.bold = true
     expect(cell.bold).to eql true
   end

   it 'add cell(0,0) and set bold to quack and check that it stays false' do
     rows = 3
     columns = 3
     paragraph = @document.paragraph(:para_001)
     table = paragraph.new_table(rows,columns)
     cell = table.cell(0,0)
     cell.bold = "quack"
     expect(cell.bold).to eql false
   end

   it 'add cell(0,0) and set align to right' do
     rows = 3
     columns = 3
     paragraph = @document.paragraph(:para_001)
     table = paragraph.new_table(rows,columns)
     cell = table.cell(0,0)
     cell.align = "right"
     expect(cell.align).to eql "right"
   end

   it 'add cell(0,0) and set align to center' do
     rows = 3
     columns = 3
     paragraph = @document.paragraph(:para_001)
     table = paragraph.new_table(rows,columns)
     cell = table.cell(0,0)
     cell.align = "center"
     expect(cell.align).to eql "center"
   end

   it 'add cell(0,0) and set align to both' do
     rows = 3
     columns = 3
     paragraph = @document.paragraph(:para_001)
     table = paragraph.new_table(rows,columns)
     cell = table.cell(0,0)
     cell.align = "both"
     expect(cell.align).to eql "both"
   end

   it 'add cell(0,0) and set align to quack and check that it stays left' do
     rows = 3
     columns = 3
     paragraph = @document.paragraph(:para_001)
     table = paragraph.new_table(rows,columns)
     cell = table.cell(0,0)
     cell.align = "quack"
     expect(cell.align).to eql "left"
   end

   it 'set create document' do
     @document.create_document()
   end
 end
