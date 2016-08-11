require 'fileutils'
require 'nokogiri'

module Wordx
  class Rels
    def initialize()
      @rels_path = "_rels/"
    end

    def create_rels(zos)

      ns = {
        "xmlns"=>"http://schemas.openxmlformats.org/package/2006/relationships",
      }
      builder = Nokogiri::XML::Builder.new do |xml|
        xml.Relationships(ns) {
          xml.Relationship "Id"=>"rId1", "Type"=>"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "Target"=>"docProps/core.xml"
          xml.Relationship "Id"=>"rId2", "Type"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "Target"=>"docProps/app.xml"
          xml.Relationship "Id"=>"rId3", "Type"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "Target"=>"word/document.xml"
        }
      end
      rels_path = @rels_path + ".rels"
      zos.put_next_entry(rels_path)
      zos.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))

    end
  end
end
