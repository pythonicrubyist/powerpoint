require 'zip/filesystem'
require 'fileutils'

module Powerpoint

  TEMPLATE_PATH = 'template/'

  class Powerpoint::Presentation

    attr_reader :pptx_path, :extract_path

    def initialize path
      @pptx_path = path
      @extract_path = path.split('.pptx').first + Time.now.strftime("%Y-%m-%d-%H%M%S")
      FileUtils.copy_entry TEMPLATE_PATH, @extract_path
    end

    def add_intro title, subtitile=nil
      intro_slide_path = "#{@extract_path}/ppt/slides/slide1.xml"
      temp = File.read intro_slide_path

      xml_title = '<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>' + title.to_s + '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
      temp.gsub!('PRESENTATION_TITLE_PACEHOLDER', xml_title)

      xml_subtitle = '<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>' + subtitile.to_s+ '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
      temp.gsub!('PRESENTATION_SUBTITLE_PACEHOLDER', xml_subtitle)

      File.open(intro_slide_path, 'w'){ |f| f << temp }
    end

    def add_textual_slide title, content=[]
      slide_path = "#{@extract_path}/ppt/slides/slide2.xml"
      temp = File.read slide_path

      title_xml = '<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>'+ title.to_s + '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
      temp.gsub!('SLIDE_TITLE_PACEHOLDER', title_xml)

      content_xml = ''
      content.each do |i|
        content_xml += '<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>' + i.to_s + '</a:t></a:r></a:p>'
      end

      temp.gsub!('CONTENT_PACEHOLDER', content_xml)

      File.open(slide_path, 'w'){ |f| f << temp }
    end

    def add_pictorial_slide title, image_path

    end

    def save
      Powerpoint.compress_pptx @extract_path, @pptx_path
    end
  end
end