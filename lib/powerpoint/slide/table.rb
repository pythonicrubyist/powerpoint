require 'zip/filesystem'
require 'fileutils'

module Powerpoint
  module Slide
    class Powerpoint::Slide::Table
      def initialize extract_path, title, table, slide_number
        template_path = "#{TEMPLATE_PATH}/ppt/slides/slide3.xml"
        xml = File.read template_path

        #set the title
        xml.gsub!("TITLE_PLACEHOLDER", title)

        #set the table
        xml.gsub! "TABLE_PLACEHOLDER", table.to_xml

        #write the result to the powerpoint
        slide_path = "#{extract_path}/ppt/slides/slide#{slide_number}.xml"
        File.open(slide_path, "w") { |f| f << xml }
      end
    end
  end
end
