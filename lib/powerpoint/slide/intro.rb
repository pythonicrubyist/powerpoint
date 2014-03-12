require 'zip/filesystem'
require 'fileutils'

module Powerpoint
  module Slide
    class Powerpoint::Slide::Intro
      def initialize extract_path, title, subtitile=nil
        template_path = "#{TEMPLATE_PATH}/ppt/slides/slide1.xml"
        xml = File.read template_path

        xml_title = '<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>' + title.to_s + '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
        xml.gsub!('PRESENTATION_TITLE_PACEHOLDER', xml_title)

        xml_subtitle = '<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>' + subtitile.to_s+ '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
        xml.gsub!('PRESENTATION_SUBTITLE_PACEHOLDER', xml_subtitle)

        intro_slide_path = "#{extract_path}/ppt/slides/slide1.xml"
        File.open(intro_slide_path, 'w'){ |f| f << xml }
      end
    end
  end
end