require 'fileutils'
require 'nokogiri'
module DocProp
  class Content
    def initialize(path = nil)

      path_rels = File.dirname(__FILE__) + "/tempdoc/rels_/" if path.nil?
      FileUtils::mkdir_p path_rels unless File.exists?(path_rels)
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
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>',''))
      end
    end

    def create_core(user = nil, description = nil, language = nil, revision = nil, subject = nil, title = nil)

      ns = {
        "xmlns:cp" => "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "xmlns:dc" => "http://purl.org/dc/elements/1.1/",
        "xmlns:dcterms" => "http://purl.org/dc/terms/"
      }
      xs = {
        "xsi:type"=>"dcterms:W3CDTF"
      }

      user = "Rail wordx/#{Wordx::VERSION}$Linux_X86_64" if user.nil?
      description = "Documented created by Rail wordx/#{Wordx::VERSION}$Linux_X86_64"
      language = "US-EN" if language.nil?
      revision = "1.0" if revision.nil?
      subject = "" if subject.nil?
      title = "" if title.nil?

      builder = Nokogiri::XML::Builder.new do |xml|
        xml[:cp].coreProperties(ns) {
          xml[:dcterms].created Time.now.strftime("%Y-%m-%dT%H:%M:%S"), xs
          xml[:dc].creator user
          xml[:dc].description description
          xml[:dc].language language
          xml[:cp].lastModifiedBy user
          xml[:dcterms].modified Time.now.strftime("%Y-%m-%dT%H:%M:%S"),xs
          xml[:cp].revision revision
          xml[:dc].subject subject
          xml[:dc].title title
        }
      end

      file_path = @docProps_path + "core.xml"
      File.delete(file_path) if File.exists?(file_path)
      File.open(file_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>',''))
      end
    end
  end
end
