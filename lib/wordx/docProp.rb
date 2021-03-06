require 'fileutils'
require 'nokogiri'
module Wordx
  class DocProp
    def initialize()
      @docProps_path = "docProps/"
    end

    def create_app(zos,pages = nil, words =nil, characters =nil, characters_with_spaces = nil, paragraphs =nil)
      pages = 1 if pages.nil?
      words = 60 if words.nil?
      characters = 180 if characters.nil?
      characters_with_spaces = 167 if characters_with_spaces.nil?
      paragraphs = 12 if paragraphs.nil?
      ns = {
        "xmlns:xsi"=>"http://www.w3.org/2001/XMLSchema-instance"
      }

      builder = Nokogiri::XML::Builder.new do |xml|
        xml.Properties(ns) {
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
      zos.put_next_entry(file_path)
      zos.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))

    end

    def create_core(zos,user = nil, description = nil, language = nil, revision = nil, subject = nil, title = nil)

      ns = {
        "xmlns:cp" => "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "xmlns:dc" => "http://purl.org/dc/elements/1.1/",
        "xmlns:dcterms" => "http://purl.org/dc/terms/",
        "xmlns:dcmitype"=>"http://purl.org/dc/dcmitype/",
        "xmlns:xsi"=>"http://www.w3.org/2001/XMLSchema-instance"
      }
      xs = {
        "xsi:type"=>"dcterms:W3CDTF"
      }

      user = "Rail wordx/#{Wordx::VERSION}$Linux_X86_64" if user.nil?
      description = "Documented created by Rail wordx/#{Wordx::VERSION}$Linux_X86_64"
      language = "en-US" if language.nil?
      revision = "1" if revision.nil?
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
      zos.put_next_entry(file_path)
      zos.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))

    end
  end
end
