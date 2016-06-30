require 'fileutils'
require 'nokogiri'
require "wordx/styles"

module Wordx
  class Word
    def initialize(path = nil)
      path = File.dirname(__FILE__) + "/tempdoc/word/" if path.nil?
      FileUtils::mkdir_p path unless File.exists?(path)
      @word_path = path
      path_rels = path + "/_rels/"
      FileUtils::mkdir_p path_rels unless File.exists?(path_rels)
      @word_rels_path = path_rels
    end

    def create_font_table(fonts = nil)

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:xsi"=>"http://www.w3.org/2001/XMLSchema-instance"
      }
      if fonts.nil?
        fonts = [{:name=>"Times New Roman",:charset=>"00",:family=>"roman",:pitch=>"variable"}]
        fonts << {:name=>"Symbol",:charset=>"02",:family=>"roman",:pitch=>"variable"}
        fonts << {:name=>"Arial",:charset=>"00",:family=>"swiss",:pitch=>"variable"}
        fonts << {:name=>"Liberation Serif",:altName =>"Times New Roman",:charset=>"01",:family=>"roman",:pitch=>"variable"}
        fonts << {:name=>"Liberation Sans",:altName =>"Arial",:charset=>"01",:family=>"swiss",:pitch=>"variable"}
        fonts << {:name=>"Courier 10 Pitch",:charset=>"01",:family=>"auto",:pitch=>"fixed"}
      end


      builder = Nokogiri::XML::Builder.new do |xml|
        xml[:w].fonts(ns) {
          fonts.each do |font_array|
              fn = {
                "w:name" =>font_array[:name]
              }
            xml[:w].font(fn) {
              xml[:w].altName("w:val"=>font_array[:altName]) unless font_array[:altName].nil?
              xml[:w].charset("w:val"=>font_array[:charset])
              xml[:w].family("w:val"=>font_array[:family])
              xml[:w].pitch("w:val"=>font_array[:pitch])
            }
          end
        }
      end
      word_path = @word_path + "fontTable.xml"
      File.delete(word_path) if File.exists?(word_path)
      File.open(word_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end
    end

    def create_rels(fonts = nil)

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:xsi"=>"http://www.w3.org/2001/XMLSchema-instance"
      }
      builder = Nokogiri::XML::Builder.new do |xml|
        xml.Relationships(ns) {
          xml.Relationship "Id"=>"rId1", "Type"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "Target"=>"styles.xml"
          xml.Relationship "Id"=>"rId2", "Type"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "Target"=>"numbering.xml"
          xml.Relationship "Id"=>"rId3", "Type"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable", "Target"=>"fontTable.xml"
          xml.Relationship "Id"=>"rId4", "Type"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "Target"=>"settings.xml"
        }
      end

      word_rels_path = @word_rels_path + "document.xml.rels"
      File.delete(word_rels_path) if File.exists?(word_rels_path)
      File.open(word_rels_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end

    end

    def create_numbering(numbering_format = nil)

      if numbering_format.nil?
        numbering_format = [{:level=>"0",:start=>"1",:style=>"Heading1",:format=>"none",:suffix=>"nothing",:justify=>"left",:tab_type=>"num",:tab_spacing=>"432",:indent=>"432",:indent_hanging=>"432"}]
        numbering_format << {:level=>"1",:start=>"1",:style=>"Heading2",:format=>"none",:suffix=>"nothing",:justify=>"left",:tab_type=>"num",:tab_spacing=>"576",:indent=>"576",:indent_hanging=>"576"}
        numbering_format << {:level=>"2",:start=>"1",:style=>"Heading3",:format=>"none",:suffix=>"nothing",:justify=>"left",:tab_type=>"num",:tab_spacing=>"720",:indent=>"720",:indent_hanging=>"720"}
        numbering_format << {:level=>"3",:start=>"1",:format=>"none",:suffix=>"nothing",:justify=>"left",:tab_type=>"num",:tab_spacing=>"864",:indent=>"864",:indent_hanging=>"864"}
        numbering_format << {:level=>"4",:start=>"1",:format=>"none",:suffix=>"nothing",:justify=>"left",:tab_type=>"num",:tab_spacing=>"1008",:indent=>"1008",:indent_hanging=>"1008"}
        numbering_format << {:level=>"5",:start=>"1",:format=>"none",:suffix=>"nothing",:justify=>"left",:tab_type=>"num",:tab_spacing=>"1152",:indent=>"1152",:indent_hanging=>"1152"}
        numbering_format << {:level=>"6",:start=>"1",:format=>"none",:suffix=>"nothing",:justify=>"left",:tab_type=>"num",:tab_spacing=>"1296",:indent=>"1296",:indent_hanging=>"1296"}
        numbering_format << {:level=>"7",:start=>"1",:format=>"none",:suffix=>"nothing",:justify=>"left",:tab_type=>"num",:tab_spacing=>"1440",:indent=>"1440",:indent_hanging=>"1440"}
        numbering_format << {:level=>"8",:start=>"1",:format=>"none",:suffix=>"nothing",:justify=>"left",:tab_type=>"num",:tab_spacing=>"1584",:indent=>"1584",:indent_hanging=>"1584"}
      end

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:o"=>"urn:schemas-microsoft-com:office:office",
        "xmlns:r"=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:v"=>"urn:schemas-microsoft-com:vml"
      }
      builder = Nokogiri::XML::Builder.new do |xml|
        xml[:w].numbering(ns) {
          abs_num = {"w:abstractNumId"=>"1"}
          xml[:w].abstractNum (abs_num) {
            numbering_format.each do |num|
              num_lvl = {"w:ilvl"=>num[:level]} unless num[:level].nil?
              xml[:w].lvl (num_lvl){
                xml[:w].start "w:val"=>num[:start] unless num[:start].nil?
                xml[:w].pStyle "w:val"=>num[:style] unless num[:style].nil?
                xml[:w].numFmt "w:val"=>num[:format] unless num[:format].nil?
                xml[:w].suff "w:val"=>num[:suffix] unless num[:suffix].nil?
                xml[:w].lvlText "w:val"=>""
                xml[:w].lvlJc "w:val"=>num[:justify] unless num[:justify].nil?
                xml[:w].pPr{
                  xml[:w].tabs{
                    tab_val = {"w:val"=>num[:tab_type], "w:pos"=>num[:tab_spacing]}
                    xml[:w].tab tab_val
                  }
                  ind_val = {"w:left"=>num[:indent],"w:hanging"=>num[:indent_hanging]}
                  xml[:w].ind(ind_val)
                }
              }
            end
          }
          num_id = {"w:numId"=>"1" }
          xml[:w].num (num_id){
            xml[:w].abstractNumId "w:val"=>"1"
          }
        }
      end

      numbering_path = @word_path + "numbering.xml"
      File.delete(numbering_path) if File.exists?(numbering_path)
      File.open(numbering_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end

    end


    def create_settings(fonts = nil)

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      }

      builder = Nokogiri::XML::Builder.new do |xml|
        xml[:w].settings(ns) {
          xml[:w].zoom "w:percent"=> "100"
          xml[:w].defaultTabStop "w:val"=>"700"
          xml[:w].compat
          xml[:w].themeFontLang "w:val"=> "", "w:eastAsia" => "", "w:bidi"=> ""
        }
      end

      word_path = @word_path + "settings.xml"
      File.delete(word_path) if File.exists?(word_path)
      File.open(word_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end

    end

    def create_document(paragraphs)

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:mc"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      }



      builder = Nokogiri::XML::Builder.new do |xml|
        xml[:w].document(ns,"mc:Ignorable"=>"w14 wp14") {
          xml[:w].body {
            paragraphs.list.each do |para_key|
              style = paragraphs.get_paragraph_style(para_key)
              xml[:w].p{
                xml[:w].pPr {
                  xml[:w].pStyle("w:val"=>style.style) unless style.nil?
                  spacing_detail = {"w:before"=>style.spacing_before,
                                    "w:after"=>style.spacing_after}
                  xml[:w].spacing (spacing_detail)
                  xml[:w].rPr
                }
                xml[:w].r{
                  xml[:w].rPr{
                    xml[:w].b("w:val"=>paragraphs.get_paragraph_bold(para_key)) unless paragraphs.nil?
                    xml[:w].bCs("w:val"=>paragraphs.get_paragraph_bold(para_key)) unless paragraphs.nil?
                  }

                  xml[:w].t (paragraphs.get_paragraph_text(para_key)) unless paragraphs.nil?
                }
              }
            end
            xml[:w].sectPr{
                xml[:w].type "w:val"=>"nextPage"
                xml[:w].pgSz "w:w"=>"11906", "w:h"=>"16838"
                xml[:w].pgMar "w:left"=>"1134", "w:right"=>"1134", "w:header"=>"0", "w:top"=>"1134", "w:footer"=>"0", "w:bottom"=>"1134", "w:gutter"=>"0"
                xml[:w].pgNumType "w:fmt"=>"decimal"
                xml[:w].formProt "w:val"=>"false"
                xml[:w].textDirection "w:val"=>"lrTb"
              }
          }
        }
      end
      doc_path = @word_path + "document.xml"
      File.delete(doc_path) if File.exists?(doc_path)
      File.open(doc_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end
    end

    def create_styles()

      doc_styles = Wordx::Styles.new

      ns = {
        "xmlns:w"=>"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:w14"=>"http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:mc"=>"http://schemas.openxmlformats.org/markup-compatibility/2006"
      }

      builder = Nokogiri::XML::Builder.new do |xml|

        xml[:w].styles(ns,"mc:Ignorable"=>"w14") {
          xml[:w].docDefaults{
            xml[:w].rPrDefault {
              xml[:w].rPr {
                xml[:w].rFonts "w:ascii"=> "Liberation Serif", "w:hAnsi"=> "Liberation Serif", "w:eastAsia"=> "Noto Sans CJK SC Regular", "w:cs"=> "FreeSans"
                xml[:w].sz "w:val"=>"24"
                xml[:w].szCs "w:val"=>"24"
                xml[:w].lang "w:val"=> "en-ZA", "w:eastAsia"=> "zh-CN", "w:bidi"=> "hi-IN"
              }
            }
            xml[:w].pPrDefault {
              xml[:w].pPr {
                xml[:w].widowControl
              }
            }
          }
          doc_styles.list.each do |doc_style_name|
            unless doc_style_name==:DefaultText
              doc_style = doc_styles.style(doc_style_name)
              style_attr = {"w:type"=>doc_style.style,"w:styleId"=>doc_style.id}
              xml[:w].style(style_attr)  {
                xml[:w].name "w:val"=>doc_style.name unless doc_style.name.nil?
                xml[:w].basedOn "w:val"=>doc_style.style_based_on unless doc_style.style_based_on.nil?
                xml[:w].next "w:val"=>doc_style.style_next unless doc_style.style_next.nil?
                xml[:w].qFormat
                xml[:w].pPr {
                  xml[:w].keepNext if doc_style.para_keep_with_next
                  spacing = {
                    "w:lineRule"=>doc_style.spacing_line_rule,
                    "w:line"=>doc_style.spacing_line,
                    "w:before"=>doc_style.spacing_before,
                     "w:after"=>doc_style.spacing_after
                  }
                  xml[:w].spacing spacing
                  xml[:w].widowControl
                }
                xml[:w].rPr {
                  font = {
                  "w:ascii"=>doc_style.font_ascii,
                    "w:hAnsi"=>doc_style.font_ansi,
                    "w:eastAsia"=>doc_style.font_lang_east_asia,
                    "w:cs"=>doc_style.font_cs
                  }
                  xml[:w].rFonts font
                  xml[:w].color "w:val"=>doc_style.font_color unless doc_style.font_color.nil?
                  xml[:w].sz "w:val"=>doc_style.font_size unless doc_style.font_size.nil?
                  xml[:w].szCs "w:val"=>doc_style.font_size unless doc_style.font_size.nil?
                  lang = {

                  "w:val"=>doc_style.font_lang_default,
                    "w:eastAsia"=>doc_style.font_lang_east_asia,
                    "w:bidi"=>doc_style.font_lang_bidi
                  }
                  xml[:w].lang
                }
              }
            end
          end
        }
      end

      doc_path = @word_path + "styles.xml"
      File.delete(doc_path) if File.exists?(doc_path)
      File.open(doc_path, "w+") do |fw|
        fw.write(builder.to_xml.sub!('<?xml version="1.0"?>','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'))
      end
    end

  end
end
