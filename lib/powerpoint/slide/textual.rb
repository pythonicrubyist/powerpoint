require 'fileutils'
require 'erb'

module Powerpoint
  module Slide
    class Textual
      include Powerpoint::Util
      
      attr_reader :title, :content

      def initialize(options={})
        require_arguments [:title, :content], options
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
        File.open("#{extract_path}/ppt/slides/slide#{index}.xml", 'w'){ |f| f << render_view('textual_slide.xml.erb') }
      end
    end
  end
end