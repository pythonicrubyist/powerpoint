require 'zip/filesystem'
require 'fileutils'

module Powerpoint
  class Powerpoint::Relationship
    def initialize extract_path, slide_number
      xml = '<Relationship Id="rId'+ (slide_number+666).to_s + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + slide_number.to_s + '.xml"/></Relationships>'
      path = "#{extract_path}/ppt/_rels/presentation.xml.rels"
      template = File.read path
      template.gsub!('</Relationships>', xml)
      File.open(path, 'w'){ |f| f << template }  
    end
  end
end
