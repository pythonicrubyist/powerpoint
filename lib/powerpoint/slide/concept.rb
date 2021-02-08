require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class Concept
      include Powerpoint::Util

      attr_reader :title, :subtitle,:subtitle_2, :logo, :image_information, :images

      def initialize(options={})
        require_arguments [:title, :subtitle,:subtitle_2, :logo, :image_information, :images], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @images = images
        slide_y = 5735412
        slide_x = 11470824
        slide_x_offset = 360588
        slide_y_offset = 1283763
        #image information = [width, height, ratio]
        # if image is taller than it is wider
        if (image_information[2] < 1)
          @image_y_scale = slide_y
          @image_x_scale = (slide_y * image_information[2]).round
          @image_x_offset = (slide_x_offset - ((slide_x - @image_x_scale) / 2)).round
          @image_y_offset = slide_y_offset
        else
          @image_y_scale = (slide_x / image_information[2]).round
          @image_x_scale = slide_x
          @image_x_offset = slide_x_offset
          @image_y_offset = (slide_y_offset + ((slide_y - @image_y_scale)/2)).round
        end
      end

      def save(extract_path, index)
        @images.each do |image|
          copy_media(extract_path, image)
        end
        
        copy_media(extract_path, logo)

        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end


      def save_rel_xml(extract_path, index)
        render_view('concept_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", index: index)
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('concept.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end