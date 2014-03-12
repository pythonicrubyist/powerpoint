require 'zip/filesystem'
require 'fileutils'

module Powerpoint

  TEMPLATE_PATH = 'template/'

  class Powerpoint::Presentation

    attr_reader :pptx_path, :extract_path

    def initialize
      @slide_count = 0
      @extract_path =  "extract_#{Time.now.strftime("%Y-%m-%d-%H%M%S")}"
      FileUtils.copy_entry TEMPLATE_PATH, @extract_path
    end

    def add_intro title, subtitile=nil
      @slide_count += 1

      intro_slide_path = "#{TEMPLATE_PATH}/ppt/slides/slide1.xml"
      template_xml = File.read intro_slide_path

      xml_title = '<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>' + title.to_s + '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
      template_xml.gsub!('PRESENTATION_TITLE_PACEHOLDER', xml_title)

      xml_subtitle = '<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>' + subtitile.to_s+ '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
      template_xml.gsub!('PRESENTATION_SUBTITLE_PACEHOLDER', xml_subtitle)

      File.open(intro_slide_path, 'w'){ |f| f << template_xml }
    end

    def add_textual_slide title, content=[]
      @slide_count += 1

      slide_template_path = "#{TEMPLATE_PATH}/ppt/slides/slide2.xml"
      template_xml = File.read slide_template_path

      title_xml = '<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>'+ title.to_s + '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
      template_xml.gsub!('SLIDE_TITLE_PACEHOLDER', title_xml)

      content_xml = ''
      content.each do |i|
        content_xml += '<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>' + i.to_s + '</a:t></a:r></a:p>'
      end

      template_xml.gsub!('CONTENT_PACEHOLDER', content_xml)

      slide_path = "#{@extract_path}/ppt/slides/slide#{@slide_count}.xml"
      File.open(slide_path, 'w'){ |f| f << template_xml }

      if @slide_count > 2
        relationship_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/></Relationships>'
        relationship_path = "#{@extract_path}/ppt/slides/_rels/slide#{@slide_count}.xml.rels"
        File.open(relationship_path, 'w'){ |f| f << relationship_xml }

        override_xml = '<Override PartName="/ppt/slides/slide' + @slide_count.to_s + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'
        content_types_path = "#{@extract_path}/[Content_Types].xml"
        override_template = File.read content_types_path
        override_template.gsub!('</Types>', override_xml)
        File.open(content_types_path, 'w'){ |f| f << override_template }

        presentaion_relationship_xml = '<Relationship Id="rId'+ (@slide_count+666).to_s + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + @slide_count.to_s + '.xml"/></Relationships>'
        presentaion_relationship_path = "#{@extract_path}/ppt/_rels/presentation.xml.rels"
        presentaion_relationship_template = File.read presentaion_relationship_path
        presentaion_relationship_template.gsub!('</Relationships>', presentaion_relationship_xml)
        File.open(presentaion_relationship_path, 'w'){ |f| f << presentaion_relationship_template }  

        presentaion_xml = '<p:sldId id="257" r:id="rId' + (@slide_count+666).to_s + '"/></p:sldIdLst>'
        presentaion_path = "#{@extract_path}/ppt/presentation.xml"
        presentaion_template = File.read presentaion_path
        presentaion_template.gsub!('</p:sldIdLst>', presentaion_xml)
        File.open(presentaion_path, 'w'){ |f| f << presentaion_template }    
      end
    end

    def add_pictorial_slide title, image_path

    end

    def save path
      @pptx_path = path
      Powerpoint.compress_pptx @extract_path, @pptx_path
      path
    end
  end
end