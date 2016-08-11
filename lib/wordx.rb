require "wordx/version"
require "wordx/content_type"
require "wordx/docProp"
require "wordx/word"
require "wordx/rels"
require "wordx/paragraphs"
require "rubygems"
require 'zip'
require 'tempfile'

module Wordx
  class Document
    def initialize()
      @styles = Wordx::Styles.new()
      @paragraphs = Wordx::Paragraphs.new()
    end

    def list_styles()
      @styles.list
    end

    def style(style)
      if style.nil?
        style = @styles.style(:DefaultText)
      else
        style = @styles.style(style)
      end
    end

    def new_paragraph(style=nil,font = nil, font_size=nil,bold=false)
      if style.nil?
        style = @styles.style(:DefaultText)
      else
        style = @styles.style(style)
      end
      font = style.font_ascii if font.nil?
      size = style.font_size() if font_size.nil?
      style.font_ascii = font
      style.font_size = font_size
      return @paragraphs.new_paragraph(style,font,font_size,bold)
    end

    def list_paragraphs()
      @paragraphs.list
    end

    def paragraph(key = nil)
      @paragraphs.paragraph(key)
    end

    def initialize_document()

    end

    def create_document()
      # Give the path of the temp file to the zip outputstream, it won't try to open it as an archive.
      compressed_filestream = Zip::OutputStream.write_buffer do |zos|

        content_type = Wordx::ContentType.new()
        content_type.create(zos)
        doc_props_app = Wordx::DocProp.new()
        doc_props_app.create_app(zos)
        doc_props_app.create_core(zos)
        rels = Wordx::Rels.new()
        rels.create_rels(zos)
        @word = Wordx::Word.new()
        @word.create_font_table(zos)
        @word.create_rels(zos)
        @word.create_settings(zos)
        @word.create_styles(zos)
        @word.create_numbering(zos)
        @word.create_document(zos,@paragraphs) unless @paragraphs.nil?


      end
      # End of the block  automatically closes the file.
      # Send it using the right mime type, with a download window and some nice file name.
      return compressed_filestream
    end
  end
end
