require 'fileutils'
require 'nokogiri'

module Rels
  class Content
    def initialize(path = nil)
      path_rels = File.dirname(__FILE__) + "/tempdoc/_rels/" if path.nil?
      FileUtils::mkdir_p path_rels unless File.exists?(path_rels)
      @rels_path = path_rels
    end

    def create_rels(rels = nil)

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
      File.delete(rels_path) if File.exists?(rels_path)
      File.open(rels_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end
    end
  end
end
