require 'zip/filesystem'
require 'fileutils'

module Powerpoint
  class Powerpoint::ContentType
    def initialize extract_path, sllide_number
      xml = '<Override PartName="/ppt/slides/slide' + sllide_number.to_s + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'
      path = "#{extract_path}/[Content_Types].xml"
      template = File.read path
      template.gsub!('</Types>', xml)
      File.open(path, 'w'){ |f| f << template }
    end
  end
end