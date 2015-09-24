require 'fileutils'
require 'erb'

module Powerpoint
  module Slide
    class Intro
      include Powerpoint::Util
      
      attr_reader :title, :subtitile

      def initialize(options={})
        require_arguments [:title, :subtitile], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
      end

      def save(index)
        save_rel_xml(index)
        save_slide_xml(index)
      end

      private

      def save_rel_xml index
        File.open("#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", 'w'){ |f| f << render_view('textual_rel.xml.erb') }
      end

      def save_slide_xml index
        File.open("#{extract_path}/ppt/slides/slide#{index}.xml", 'w'){ |f| f << render_view('intro_slide.xml.erb') }
      end
    end
  end
end