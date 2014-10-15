
module Powerpoint
  module Utils
    class Table
      def initialize matrix, settings = {}
        raise ArgumentError.new("no rows") if matrix.first.nil?
        raise ArgumentError.new("no cols") if matrix.first.first.nil?

        #actual rows and cols
        @matrix   = matrix

        #number of cols...used for the tblgrid section
        @ncols    = matrix.first.length

        #column width
        #TODO: specify individual rows
        #TODO: human readable widths
        @width    = (settings[:width]    or "1645920")

        #row height
        #TODO: human readable heights
        @height   = (settings[:height]   or "370840")

        #table style by microsoft id
        @style    = name_to_guid (settings[:style] or :medium_style_2_accent_1)

        #boolean...first row is a header
        #TODO: actually covert this to a ruby bool
        @firstrow = (settings[:firstrow] or "1")

        #boolean...row color bands
        #TODO: actually convert this to a ruby bool
        @bandrow  = (settings[:bandrow]  or "1")
        
      end

      #setters
      def width=(width)
        @width = width
      end

      def style=(style)
        @style = style
      end

      def firstrow=(firstrow)
        @firstrow = firstrow
      end

      def bandrow=(bandrow)
        @bandrow = bandrow
      end

      def to_xml
		%{
			<a:tbl>
				<a:tblPr firstRow="#{@firstrow}" bandRow="">
					<a:tableStyleId>{#{@style}}</a:tableStyleId>
				</a:tblPr>
				<a:tblGrid>
					#{ 
                        "<a:gridCol w='#{@width}' />"*@ncols 
                    }
				</a:tblGrid>
				#{
					@matrix.map do |row|
						%{
							<a:tr h="#{@height}">
								#{
									row.map do |col|
										%{
											<a:tc>
												<a:txBody>
													<a:bodyPr />
													<a:lstStyle />
													<a:p>
														<a:r>
															<a:rPr lang="en-US" dirty="0" smtClean="0" />
															<a:t>#{col}</a:t>
														</a:r>
														<a:endParaRPr lang="en-US" dirty="0" />
													</a:p>
												</a:txBody>
												<a:tcPr />
											</a:tc>
										}
									end.join
								}
							</a:tr>
						}
					end.join
				}
			</a:tbl>
		}
      end

      private
      def name_to_guid(name)
        {
          no_style_no_grid:"2d5abb2605874c30899992f81fd0307c",
          theme_style_1_accent_1:"3c2ffa5d87b4456a98211d502468cf0f",
          theme_style_1_accent_2:"284e427a3d554303bf806455036e1de7",
          theme_style_1_accent_3:"69c7853c536d4a76a0aedd22124d55a5",
          theme_style_1_accent_4:"775dcb029bb847fd890785c794f793ba",
          theme_style_1_accent_5:"35758fb79ac545528a53c91805e547fa",
          theme_style_1_accent_6:"08fb837dc8274efaa0574d05807e0f7c",
          no_style_table_grid:"5940675ab579460e94d154222c63f5da",
          theme_style_2_accent_1:"d113a9d29d6b4929aa2df23b5ee8cbe7",
          theme_style_2_accent_2:"18603fdce32a4ab5989c0864c3ead2b8",
          theme_style_2_accent_3:"306799f8075e4a3aa7f67fbc6576f1a4",
          theme_style_2_accent_4:"e269d01ebc324049b4635c60d7b0ccd2",
          theme_style_2_accent_5:"327f97bbc8334fb7bde53f7075034690",
          theme_style_2_accent_6:"638b18551b754fbe930c398ba8c253c6",
          light_style_1:"9d7b26c541074fecaedc1716b250a1ef",
          light_style_1_accent_1:"3b4b98b060ac42c2afa5b58cd77fa1e5",
          light_style_1_accent_2:"0e3fde45af774b5c971549d594bdf05e",
          light_style_1_accent_3:"c083e6e3fa7d4d7ba595ef9225afea82",
          light_style_1_accent_4:"d27102a983104765a935a1911b00ca55",
          light_style_1_accent_5:"5fd0f851ec5a4d38b0ad8093ec10f338",
          light_style_1_accent_6:"68d230f3cf8048598ce7a43ee81993b5",
          light_style_2:"7e9639d4e3e24d3492845a2195b3d0d7",
          light_style_2_accent_1:"69012ecd51fc41f1aa8d1b2483cd663e",
          light_style_2_accent_2:"72833802fef14c798d5d14cf1eaf98d9",
          light_style_2_accent_3:"f2de63d5997a4646a3774702673a728d",
          light_style_2_accent_4:"17292a2ef33343fb96215cbbe7fdcdcb",
          light_style_2_accent_5:"5a111915be364e01a7e504b1672ead32",
          light_style_2_accent_6:"912c8c8551f0491e97743900afef0fd7",
          light_style_3:"616da210fb5b4158b5e0feb733f419ba",
          light_style_3_accent_1:"bc89ef968cea46ff86c44ce0e7609802",
          light_style_3_accent_2:"5da37d80643444d0a0281b22a696006f",
          light_style_3_accent_3:"8799b23bec834686b30a512413b5e67a",
          light_style_3_accent_4:"ed083ae646fa4a598fb09f97eb10719f",
          light_style_3_accent_5:"bdbed56947974df1a0f46aab3cd982d8",
          light_style_3_accent_6:"e8b1032cea384f05ba0d38afffc7bed3",
          medium_style_1:"793d81cf94f2401aba5792f5a7b2d0c5",
          medium_style_1_accent_1:"b301b821a1ff4177aee776d212191a09",
          medium_style_1_accent_2:"9dcaf9ed07dc4a118d7f57b35c25682e",
          medium_style_1_accent_3:"1fecb4d8db024dc6a0a24f2ebae1dc90",
          medium_style_1_accent_4:"1e17193346194e119a3ff7608df75f80",
          medium_style_1_accent_5:"fabfcf233b69468fb69f88f6de6a72f2",
          medium_style_1_accent_6:"10a1b5d59b994c35a422299274c87663",
          medium_style_2:"073a0daa6af343ab8588cec1d06c72b9",
          medium_style_2_accent_1:"5c22544a7ee64342b04885bdc9fd1c3a",
          medium_style_2_accent_2:"21e4aea48dfa4a8987eb49c32662afe0",
          medium_style_2_accent_3:"f5ab1c696edb4ff4983f18bd219ef322",
          medium_style_2_accent_4:"00a15c55851742aab614e9b94910e393",
          medium_style_2_accent_5:"7df18680e05441ad8bc1d1aef772440d",
          medium_style_2_accent_6:"93296810a8854be3a3e76d5beea58f35",
          medium_style_3:"8ec20e35a1764012bc5e935cfff8708e",
          medium_style_3_accent_1:"6e25e6493f164e02a73319d2cdbf48f0",
          medium_style_3_accent_2:"85be263cdbd74a20bb59aab30acaa65a",
          medium_style_3_accent_3:"eb344d849afb497ea393dc336ba19d2e",
          medium_style_3_accent_4:"eb9631b578f241c9869b9f39066f8104",
          medium_style_3_accent_5:"74c1a8a3306a4eb7a6b14f7e0eb9c5d6",
          medium_style_3_accent_6:"2a488322f2ba4b5b97480d474271808f",
          medium_style_4:"d7ac3ccac7974891be02d94e43425b78",
          medium_style_4_accent_1:"69cf1ab219764502bf363ff5ea218861",
          medium_style_4_accent_2:"8a107856555442fbb03e39f5dbc370ba",
          medium_style_4_accent_3:"0505e3ef67ea436b97b20124c06ebd24",
          medium_style_4_accent_4:"c4b1156a380e4f78bdf5a606a8083bf9",
          medium_style_4_accent_5:"22838bef8bb2449884a7c5851f593df1",
          medium_style_4_accent_6:"16d9f66e5eb9488286fbdcbf35e3c3e4",
          dark_style_1:"e8034e787f5d4c2eb375fc64b27bc917",
          dark_style_1_accent_1:"125e5076381047ddb79f674d7ad40c01",
          dark_style_1_accent_2:"37ce84f328c3443e9e9699cf82512b78",
          dark_style_1_accent_3:"d03447bb5d67496b8e87e561075ad55c",
          dark_style_1_accent_4:"e929f9f44a8f4326a1b422849713ddab",
          dark_style_1_accent_5:"8fd4443ef9894fc4a0c8d5a2af1f390b",
          dark_style_1_accent_6:"af6068537671496a8e4fdf71f8ec918b",
          dark_style_2:"5202b0cafc5444968bca5ef66a818d29",
          dark_style_2_accent_1_accent_2:"0660b408b3cf4a9485fc2b1e0a45f4a2",
          dark_style_2_accent_3_accent_4:"91ebbbccdad2459cbe2ef6de35cf9a28",
          dark_style_2_accent_5_accent_6:"46f890a928074ebbb81db2aa78ec7f39"
        }[name]
      end
    end
  end
end
