require 'zip/filesystem'
require 'fileutils'

module Powerpoint
  class Powerpoint::Meta
    def initialize extract_path, sllide_number
      xml = '<p:sldId id="257" r:id="rId' + (sllide_number+666).to_s + '"/></p:sldIdLst>'
      path = "#{extract_path}/ppt/presentation.xml"
      template = File.read path
      template.gsub!('</p:sldIdLst>', xml)
      File.open(path, 'w'){ |f| f << template } 
    end
  end
end