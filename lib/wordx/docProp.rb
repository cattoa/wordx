require 'fileutils'
require 'nokogiri'
module DocProp
  class Content
    def initialize(path = nil)
      path = File.dirname(__FILE__) + "/tempdoc/docProps/" if path.nil?
      FileUtils::mkdir_p path unless File.exists?(path)
      @docProps_path = path
    end

    def create_app(pages = nil, words =nil, characters =nil, characters_with_spaces = nil, paragraphs =nil)
      pages = 1 if pages.nil?
      words = 1 if words.nil?
      characters = 1 if characters.nil?
      characters_with_spaces = 1 if characters_with_spaces.nil?
      paragraphs = 1 if paragraphs.nil?

      builder = Nokogiri::XML::Builder.new do |xml|
        xml.Properties {
          xml.Template{}
          xml.TotalTime "4"
          xml.Application "Rail wordx/#{Wordx::VERSION}$Linux_X86_64"
          xml.Pages pages
          xml.Words words
          xml.Characters characters
          xml.CharactersWithSpaces characters_with_spaces
          xml.Paragraphs paragraphs
        }
      end

      file_path = @docProps_path + "app.xml"
      File.delete(file_path) if File.exists?(file_path)
      File.open(file_path, "w+") do |fw|
        fw.write(builder.to_xml)
      end
    end

    def create_core()

      ns = {
        "xmlns:cp" => "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "xmlns:dc" => "http://purl.org/dc/elements/1.1/",
        "xmlns:dcterms" => "http://purl.org/dc/terms/"
      }
      xs = {
        "xsi:type"=>"dcterms:W3CDTF"
      }

      builder = Nokogiri::XML::Builder.new do |xml|
        xml[:cp].coreProperties(ns) {
          xml[:dcterms].created(xs) ( Time.now.strftime("%Y-%m-%dT%H:%M:%S"))
          xml[:dc].creator
          xml[:dc].description
          xml[:dc].language
          xml[:cp].lastModifiedBy
          xml[:dcterms].modified Time.now.strftime("%Y-%m-%dT%H:%M:%S")
          xml[:cp].revision
          xml[:dc].subject
          xml[:dc].title
        }
      end

      file_path = @docProps_path + "core.xml"
      File.delete(file_path) if File.exists?(file_path)
      File.open(file_path, "w+") do |fw|
        fw.write(builder.to_xml)
      end
    end
  end
end
