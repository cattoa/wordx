require 'fileutils'
require 'nokogiri'
module Wordx
  module Document
    class Styles
      def self.initialize_style
        path = File.expand_path(File.dirname(__FILE__) + '/styles/')
        Dir.open(path) do |dir|
          dir.each do |file|
            _register_style(path + '/' + file) if file.include?('.yml')
          end
        end

      end

      def self._register_style(style_file_path)
        style = Style.new(YAML.load(IO.read(File.expand_path(style_file_path))))
        if !defined? @@_styles_list
          @@_styles_list = {}
        end
        @@_styles_list[style.name.to_sym] = style
      end

      def self.list
        @@_styles_list.keys
      end

      protected

      def self.method_missing(method, *args)
        if @@_styles_list.keys.include?(method)
          return @@_themes_list[method]
        end
      end

      class Style
        attr_accessor :name, :style, :id, :font_ascii, :font_ansi, :font_east_asia, :font_cs, :font_size, :font_color, :font_lang_default, :font_lang_east_asia,
        :font_lang_bidi, :style_based_on, :style_next, :spacing_line_rule, :spacing_line, :spacing_before, :spacing_after, :indent_left,
        :indent_hanging, :outline_level, :para_keep_with_next, :para_supress_line_numers



        def initialize(style_hash)
          @name = style_hash['name']
          @id = style_hash['id']
          @style = style_hash['style']
        end


        if style_hash.keys.include?('font_ascii')
          @font_ascii = style_hash['font_ascii']
        end


        if style_hash.keys.include?('font_ansi')
          @font_ansi = style_hash['font_ansi']
        end


        if style_hash.keys.include?('font_east_asia')
          @font_east_asia = style_hash['font_east_asia']
        end


        if style_hash.keys.include?('font_cs')
          @font_cs = style_hash['font_cs']
        end


        if style_hash.keys.include?('font_size')
          @font_size = style_hash['font_size']
        end


        if style_hash.keys.include?('font_color')
          @font_color = style_hash['font_color']
        end
        if style_hash.keys.include?('font_lang_default')
          @font_lang_default = style_hash['font_lang_default']
        end
        if style_hash.keys.include?('font_ansi')
          @font_ansi = style_hash['font_ansi']
        end
        if style_hash.keys.include?('font_ansi')
          @font_ansi = style_hash['font_ansi']
        end
        if style_hash.keys.include?('font_ansi')
          @font_ansi = style_hash['font_ansi']
        end
        if style_hash.keys.include?('font_ansi')
          @font_ansi = style_hash['font_ansi']
        end
        if style_hash.keys.include?('font_ansi')
          @font_ansi = style_hash['font_ansi']
        end


    end
  end
end
