require 'zip/filesystem'
require 'fileutils'

module Powerpoint
  module Slide
    class Powerpoint::Slide::Textpic
      def initialize extract_path, title, image_path, slide_number, image_coords={}, content=[], content_coords={}
      	
      	image_name = File.basename(image_path)
      	FileUtils.copy_file(image_path, "#{extract_path}/ppt/media/#{image_name}")

      	image_coords_tag = (image_coords.empty?) ? '' : '<p:spPr>
									            <a:xfrm>
									                <a:off x="' + image_coords[:x].to_s + '" y="' + image_coords[:y].to_s + '" />
									                <a:ext cx="' + image_coords[:cx].to_s + '" cy="' + image_coords[:cy].to_s + '" />
									            </a:xfrm>
									        </p:spPr>'

				content_coords_tag = (content_coords.empty?) ? '' : 
									            '<a:xfrm>
									                <a:off x="' + content_coords[:x].to_s + '" y="' + content_coords[:y].to_s + '" />
									                <a:ext cx="' + content_coords[:cx].to_s + '" cy="' + content_coords[:cy].to_s + '" />
									            </a:xfrm>'

        rel_xml =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
					<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
					    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' + image_name + '" />
					    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml" />
					</Relationships>'

        rel_path = "#{extract_path}/ppt/slides/_rels/slide#{slide_number}.xml.rels"
        File.open(rel_path, 'w'){ |f| f << rel_xml }

        content_xml = ''

        if content.is_a?(Hash)
        	content.each do |key, text|
        		content_xml += '<a:p>
        											<a:r>
	          										<a:rPr lang="en-US" b="1" dirty="0" smtClean="0"/>
	          										<a:t>' + key.capitalize.to_s + ':</a:t>
	          									</a:r>
	          									<a:r>
	          										<a:rPr lang="en-US" dirty="0" smtClean="0"/>
	          										<a:t> ' + text.to_s + '</a:t>
	          									</a:r>
	          									<a:endParaRPr lang="en-US" dirty="0"/>
	          								</a:p>'
	        end
        elsif content.is_a?(Array)
	        content.each do |text|
	          content_xml += '<a:p>
	          									<a:r>
	          										<a:rPr lang="en-US" dirty="0" smtClean="0"/>
	          										<a:t>' + text.to_s + '</a:t>
	          									</a:r>
	          									<a:endParaRPr lang="en-US" dirty="0"/>
	          								</a:p>'
	      	end
	      elsif content.is_a?(String)
	      	content_xml += '<a:p>
	          									<a:r>
	          										<a:rPr lang="en-US" dirty="0" smtClean="0"/>
	          										<a:t>' + content.to_s + '</a:t>
	          									</a:r>
	          									<a:endParaRPr lang="en-US" dirty="0"/>
	          								</a:p>'
        end

				slide_xml =  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
					<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
					    <p:cSld>
					        <p:spTree>
					            <p:nvGrpSpPr>
					                <p:cNvPr id="1" name="" />
					                <p:cNvGrpSpPr />
					                <p:nvPr />
					            </p:nvGrpSpPr>
					            <p:grpSpPr>
					                <a:xfrm>
					                    <a:off x="0" y="0" />
					                    <a:ext cx="0" cy="0" />
					                    <a:chOff x="0" y="0" />
					                    <a:chExt cx="0" cy="0" />
					                </a:xfrm>
					            </p:grpSpPr>
					            <p:sp>
					                <p:nvSpPr>
					                    <p:cNvPr id="2" name="Title 1" />
					                    <p:cNvSpPr>
					                        <a:spLocks noGrp="1" />
					                    </p:cNvSpPr>
					                    <p:nvPr>
					                        <p:ph type="title" />
					                    </p:nvPr>
					                </p:nvSpPr>
					                <p:spPr />
					                <p:txBody>
					                    <a:bodyPr />
					                    <a:lstStyle />
					                    <a:p>
					                        <a:r>
					                            <a:rPr lang="en-US" dirty="0" smtClean="0" />
					                            <a:t>' + title + '</a:t>
					                        </a:r>
					                        <a:endParaRPr lang="en-US" dirty="0" />
					                    </a:p>
					                </p:txBody>
					            </p:sp>
					            <p:pic>
					                <p:nvPicPr>
					                    <p:cNvPr id="4" name="Content Placeholder 3" />
					                    <p:cNvPicPr>
					                        <a:picLocks noGrp="1" noChangeAspect="1" />
					                    </p:cNvPicPr>
					                    <p:nvPr>
					                        <p:ph idx="1" />
					                    </p:nvPr>
					                </p:nvPicPr>
					                <p:blipFill>
					                    <a:blip r:embed="rId2">
					                        <a:extLst>
					                            <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
					                                <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0" />
					                            </a:ext>
					                        </a:extLst>
					                    </a:blip>
					                    <a:stretch>
					                        <a:fillRect />
					                    </a:stretch>
					                </p:blipFill>' + image_coords_tag +
					            '</p:pic>
					            <p:sp>
					                <p:nvSpPr>
					                    <p:cNvPr id="5" name="Content 1" />
					                    <p:cNvSpPr txBox="1"/>
					                    <p:nvPr/>
					                </p:nvSpPr>
					                <p:spPr>' + content_coords_tag +
							            	'<a:prstGeom prst="rect">
															<a:avLst/>
														</a:prstGeom>
														<a:noFill/>
					                </p:spPr>
					                <p:txBody> 
					                	<a:bodyPr wrap="square" rtlCol="0">
														<a:spAutoFit/>
														</a:bodyPr>
														<a:lstStyle/> '+ content_xml + '</p:txBody>
					            </p:sp>
					        </p:spTree>
					        <p:extLst>
					            <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
					                <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1474098577" />
					            </p:ext>
					        </p:extLst>
					    </p:cSld>
					    <p:clrMapOvr>
					        <a:masterClrMapping />
					    </p:clrMapOvr>
					</p:sld>'

        slide_path = "#{extract_path}/ppt/slides/slide#{slide_number}.xml"
        File.open(slide_path, 'w'){ |f| f << slide_xml }        
      end
    end
  end
end