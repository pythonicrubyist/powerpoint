require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class Gallery
      include Powerpoint::Util

      attr_reader :title, :subtitle, :images

      def initialize(options={})
        require_arguments [:title, :subtitle, :images], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @images = images
      end

      def save(extract_path, index)
        @images.each do |image|
          copy_media(extract_path, image[0])
        end

        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end


      def save_rel_xml(extract_path, index)
        render_view('gallery_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", index: index)
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('gallery.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end