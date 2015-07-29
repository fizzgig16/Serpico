# encoding: ASCII-8BIT

# Responsible for parsing Lua content in Serpico templates

require 'rubygems'
require 'nokogiri'
require 'rlua'

$W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def run_lua(document, report_xml)
	###############################
	# « - begin Lua block
	# » - end Lua block
	
	# Find any labels we might need later
	#build_label_array(document)
	
	lua = SerpicoLua.new(document, report_xml)
		
	#puts "Doc before: #{document}\n\n"
	#puts "Report: #{report_xml}\n\n"
	
	document.scan(/«(.*?)»/).each do |lua_block|
		lua_block = lua_block[0]	# Only get the capture group
		
		# Replace paragraph ends with newlines
		lua_block.gsub!(/<\/w:p>/, "\n")
		
		# Strip Word markup
		lua_block.gsub!(/<\/?w:.*?\/?>/, "")
		
		#puts "Found Lua block: #{lua_block}"
		lua.run_lua_block(lua_block)
	end
	
	# After our processing is done, replace document with whatever is in the Lua state
	document = lua.get_document()
	#puts document
	
	return document
end

class SerpicoLua
	def initialize(doc, report_xml)
		# Start the Lua engine
		@state = Lua::State.new()
		
		# Store document in state so we can work with it
		@state.document = doc
		
		# Build a nokogiri document from the doc we passed in so we only parse it once
		@noko = Nokogiri::XML(@state.document)
		
		# Store report XML
		report_xml = report_xml
		
		# Build the nokogiri doc for the report XML too
		@noko_report = Nokogiri::XML(report_xml)
		
		# Now set up our tables so that we can call stuff
		create_lua_tables()
	end

	def run_lua_block(lua_block)
		# Clean up stupid stuff like smart quotes and double-dashes
		
		puts "Calling Lua block"
		
		clean_block = lua_block
		clean_block.gsub!("“", "\"")
		clean_block.gsub!("”", "\"")
		clean_block.gsub!("–", "--")
		
		# Run ze code
		@state.__eval(clean_block)
	end

	def get_document()
		return @state.document
	end
	
	def create_lua_tables()
		#@state.__load_stdlib :math, :string, :table	# Only load bare minimum tables
		@state.__load_stdlib :all
		
		# Create tables we can call from lua to do stuff
		@state.Label = 
		{
			# Label:Replace(labelname, value)
			'Replace' => lambda { |this, labelname, value| lua_label_replace(this, labelname, value) },
			# Label:ForeColor(labelname, rgbstring) - RGB string will be something like 0070c0
			'ForeColor' => lambda { |this, labelname, rgbstring| lua_label_forecolor(this, labelname, rgbstring) }
		}
		@state.Cell = 
		{
			# Cell:BackColor(labelname, rgbstring)
			'BackColor' => lambda { |this, labelname, rgbstring| lua_cell_backcolor(this, labelname, rgbstring) }
		}
		@state.Row = 
		{
			# Row:BackColor(labelname, rgbstring)
			'BackColor' => lambda { |this, labelname, rgbstring| lua_row_backcolor(this, labelname, rgbstring) }
		}
		@state.ReportContent = 
		{
			# ReportContent:GetShortCompanyName()
			'GetShortCompanyName' => lambda { |this| lua_reportcontent_getshortcompanyname(this) }
		}
		@state.Findings = 
		{
			# Findings:GetAllFindings() - returns a table of Findings
			'GetAllFindings' => lambda { |this| lua_findings_getallfindings(this) }
		}
			

		@state.Finding = 
		{
			# Finding:GetTitle(finding)
			'GetTitle' => lambda { |this| return lua_finding_gettitle(this) },
			'GetID' => lambda { |this| return lua_finding_getid(this) }
		}

		@state.WordTable =
		{
			# WordTable:Create(columns, header_row_count, body_row_count)
			'Create' => lambda { |this, columns, header_row_count, body_row_count| lua_wordtable_create(this, columns, header_row_count, body_row_count) }
		}
		
	end

	# Gets the label complete with delimiters
	def get_full_label(label)
		return "¤#{label}¤"
	end

	#### Label functions ####
	# Replaces an entire label with another value. This destroys the label, so be careful when you call it!
	def lua_label_replace(this, labelname, value)
		full_label = get_full_label(labelname)
		@state.document = @state.document.gsub!(/#{full_label}/, value)
	end

	# Sets the foreground color of the label text
	def lua_label_forecolor(this, labelname, rgbstring)
		full_label = get_full_label(labelname)
		
		# Need to see if our label is immediately preceded by <w:t>; if so, means we have our own text block we can colorize, or at least star with
		# If not, need to end the block we're currently in and start a new one with the color of our choosing
		@noko.xpath("//w:t[contains(text(), \"#{full_label}\")]", 'w' => $W_URI).each do |text_node|
			#puts "Found #{text_node}"
			run = text_node.xpath("../../w:r", 'w' => $W_URI).first
			run_props = text_node.xpath("../w:rPr", 'w' => $W_URI).first			

			if (text_node.content == full_label)
				# Text is by itself, we can apply the color directly
				run_props.add_child(Nokogiri::XML::Node.new("color w:val=\"#{rgbstring}\"", @noko))
			else
				# Split out the text before and after the label, make sure to keep the run properties the same, and apply the new properties to only the label
				arr_fragments = text_node.content.force_encoding("ASCII-8BIT").split(full_label, 2)
				new_run = run
				new_text_node = text_node
				
				#puts "Frags: #{arr_fragments}"
				while (arr_fragments.length >= 2) # Means we have more than one label to deal with
					before = arr_fragments[0]
					after = arr_fragments[1]
					new_text_node.content = before
					
					# Now add a new run for the label
					new_run = new_run.add_next_sibling(Nokogiri::XML::Node.new("r", @noko))
					new_run_props = new_run.add_child(Nokogiri::XML::Node.new("rPr", @noko))
					new_run_props.add_child(Nokogiri::XML::Node.new("color w:val=\"#{rgbstring}\"", @noko))
					new_text_node = new_run.add_child(Nokogiri::XML::Node.new("t", @noko))
					new_text_node.content = full_label
					
					# Now the next run
					if (after != "")
						new_run = new_run.add_next_sibling(Nokogiri::XML::Node.new("r", @noko))
						new_run_props = new_run.add_child(Nokogiri::XML::Node.new("rPr", @noko))
						new_text_node = new_run.add_child(Nokogiri::XML::Node.new("t", @noko))
		
						arr_fragments = after.split(full_label, 2)
						if (arr_fragments.length < 2)
							# Add the last part and call it a day
							new_text_node.content = arr_fragments[0]
							break
						end
					else
						break
					end
				end
			end
			
			#puts "Content: #{@noko}\n\n"
			@state.document = @noko.to_s()
							
			#parent = text_node.xpath("..", 'w' => $W_URI)
			#puts "My parent is #{parent}"
		end
		
		@state.document = @state.document.gsub!(/#{full_label}/, "<w:color w:val=\"#{rgbstring}\"/>#{full_label}")
	end
			
	#### Cell functions ####
	def lua_cell_backcolor(this, labelname, rgbstring)
		full_label = get_full_label(labelname)
		
		@noko.xpath("//w:tc/w:p/w:r/w:t[contains(text(), \"#{full_label}\")]", 'w' => $W_URI).each do |cell_node|
			#puts "Found #{cell_node}"
			shading_node = cell_node.xpath("../../../w:tcPr/w:shd", 'w' => $W_URI)
			if (shading_node.empty?)
				prop_node = cell_node.xpath("../../../w:tcPr", 'w' => $W_URI)
				#<w:shd w:fill="FF0000" w:val="clear"/>
				prop_node.add_child(Nokogiri::XML::Node.new("shd w:fill=\"#{rgbstring}\" w:val=\"clear\"", @noko))
			else
				#puts "Found shading node"
				shading_node = shading_node.first
				shading_node["w:fill"] = rgbstring
			end
			
			#puts "Content: #{@noko}\n\n"
			@state.document = @noko.to_s()
		end
	end
	
	#### Row functions ####
	def lua_row_backcolor(this, labelname, rgbstring)
		full_label = get_full_label(labelname)
		
		@noko.xpath("//w:tr/w:tc/w:p/w:r/w:t[contains(text(), \"#{full_label}\")]", 'w' => $W_URI).each do |row_node|
			puts "Found #{row_node}"
			cell_nodes = row_node.xpath("../../../../w:tc", 'w' => $W_URI)
			cell_nodes.each do |cell_node|
				#puts "node"
				shading_node = cell_node.xpath("./w:tcPr/w:shd", 'w' => $W_URI)
				if (shading_node.empty?)
					#puts "No shading node"
					prop_node = cell_node.xpath("../../../w:tcPr", 'w' => $W_URI)
					#<w:shd w:fill="FF0000" w:val="clear"/>
					prop_node.add_child(Nokogiri::XML::Node.new("shd w:fill=\"#{rgbstring}\" w:val=\"clear\"", @noko))
				else
					#puts "Found shading node"
					shading_node = shading_node.first
					shading_node["w:fill"] = rgbstring
				end
			end
			
			#puts "Content: #{@noko}\n\n"
			@state.document = @noko.to_s()
		end
	end
	
	#### ReportContent functions ####
	def lua_reportcontent_getshortcompanyname(this)
		return @noko_report.xpath("/report/reports/short_company_name").first.content
	end
	
	#### Findings functions ####
	
	# Returns a table of Finding(s)
	def lua_findings_getallfindings(this)
		#puts "GetAllFindings()"
		output = []
		findings = @noko_report.xpath("/report/findings_list/findings")
		findings.each do |finding|
			# Create a table
			#puts "Creating finding table"
			table = {}
			
			# Functions
			table["GetTitle"] = lambda { |table| lua_finding_gettitle(table) }
			table["GetID"] = lambda { |table| lua_finding_getid(table) }
			
			# Attributes
			table["id"] = finding.xpath("id").first.content
			table["title"] = finding.xpath("title").first.content
			output << table
		end
		
		return output
	end
	
	#### Finding functions ####
	def lua_finding_gettitle(this)
		#puts "GetTitle returning #{this["title"]}"
		return this["title"]
	end
	
	def lua_finding_getid(this)
		#puts "GetTitle returning #{this["title"]}"
		return this["id"]
	end
	
	#### WordTable functions ####
	def lua_wordtable_create(this, columns, header_row_count, body_row_count)

		# Convert strings to ints
		#columns = columns.to_i()
		#header_row_count = header_row_count.to_i()
		#body_row_count = body_row_count.to_i()

		# Creates a table in word where the code is
		total_width = 9683

		table = @noko.root.add_child(Nokogiri::XML::Node.new("tbl", @noko))
		table_params = table.add_child(Nokogiri::XML::Node.new("tblPr", @noko))
		table_params.add_child(Nokogiri::XML::Node.new("tblW w:type=\"dxa\" w:w=\"#{total_width}\"", @noko))
		table_params.add_child(Nokogiri::XML::Node.new("jc w:val=\"left\"", @noko))
		table_params.add_child(Nokogiri::XML::Node.new("tblW w:type=\"dxa\" w:w=\"55\"", @noko))
		table_borders = table_params.add_child(Nokogiri::XML::Node.new("tblBorders", @noko))
		table_borders.add_child(Nokogiri::XML::Node.new("top w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
		table_borders.add_child(Nokogiri::XML::Node.new("left w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
		table_borders.add_child(Nokogiri::XML::Node.new("bottom w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
		table_borders.add_child(Nokogiri::XML::Node.new("insideH w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
		table_borders.add_child(Nokogiri::XML::Node.new("right w:val=\"nil\"", @noko))
		table_borders.add_child(Nokogiri::XML::Node.new("insideV w:val=\"nil\"", @noko))
		cell_margins = table_params.add_child(Nokogiri::XML::Node.new("tblCellMar", @noko))
		cell_margins.add_child(Nokogiri::XML::Node.new("top w:type=\"dxa\" w:w=\"55\"", @noko))
		cell_margins.add_child(Nokogiri::XML::Node.new("left w:type=\"dxa\" w:w=\"55\"", @noko))
		cell_margins.add_child(Nokogiri::XML::Node.new("bottom w:type=\"dxa\" w:w=\"55\"", @noko))
		cell_margins.add_child(Nokogiri::XML::Node.new("right w:type=\"dxa\" w:w=\"55\"", @noko))
		table_grid = table.add_child(Nokogiri::XML::Node.new("tblGrid", @noko))

		for cols in 1..columns
			table_grid.add_child(Nokogiri::XML::Node.new("gridCol w:w=\"#{total_width / columns}\"", @noko))
		end
	
		# Now the rows, header first
		for header_rows in 1..header_row_count
			header_row = table.add_child(Nokogiri::XML::Node.new("tr", @noko))
			header_props = header_row.add_child(Nokogiri::XML::Node.new("trPr", @noko))
			header_props.add_child(Nokogiri::XML::Node.new("cantSplit w:val=\"false\"", @noko))
			
			# Add the cells
			for cols in 1..columns
				header_cell = header_row.add_child(Nokogiri::XML::Node.new("tc", @noko))
				header_cell_props = header_cell.add_child(Nokogiri::XML::Node.new("tcPr", @noko))
				header_cell_props.add_child(Nokogiri::XML::Node.new("tcW w:type=\"dxa\" w:w=\"#{total_width / columns}\"", @noko))
				header_cell_borders = header_cell_props.add_child(Nokogiri::XML::Node.new("tcBorders", @noko))
				header_cell_borders.add_child(Nokogiri::XML::Node.new("top w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				header_cell_borders.add_child(Nokogiri::XML::Node.new("left w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				header_cell_borders.add_child(Nokogiri::XML::Node.new("bottom w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				header_cell_borders.add_child(Nokogiri::XML::Node.new("right w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				header_cell_props.add_child(Nokogiri::XML::Node.new("shd w:fill=\"000000\" w:val=\"clear\"", @noko))
				header_cell_margins = header_cell_props.add_child(Nokogiri::XML::Node.new("tcMar", @noko))
				header_cell_margins.add_child(Nokogiri::XML::Node.new("left w:type=\"dxa\" w:w=\"55\"", @noko))
				header_paragraph = header_cell.add_child(Nokogiri::XML::Node.new("p", @noko))
				header_paragraph_props = header_paragraph.add_child(Nokogiri::XML::Node.new("pPr", @noko))
				header_paragraph_props.add_child(Nokogiri::XML::Node.new("pStyle w:val=\"style20\"", @noko))
				header_paragraph_props.add_child(Nokogiri::XML::Node.new("rPr", @noko))
				header_paragraph_run = header_paragraph.add_child(Nokogiri::XML::Node.new("r", @noko))
				header_paragraph_run.add_child(Nokogiri::XML::Node.new("rPr", @noko))
			end
		end
		
		# And the body rows
		for body_rows in 1..body_row_count
			body_row = table.add_child(Nokogiri::XML::Node.new("tr", @noko))
			body_props = body_row.add_child(Nokogiri::XML::Node.new("trPr", @noko))
			body_props.add_child(Nokogiri::XML::Node.new("cantSplit w:val=\"false\"", @noko))
			
			# Add the cells
			for cols in 1..columns
				body_cell = body_row.add_child(Nokogiri::XML::Node.new("tc", @noko))
				body_cell_props = body_cell.add_child(Nokogiri::XML::Node.new("tcPr", @noko))
				body_cell_props.add_child(Nokogiri::XML::Node.new("tcW w:type=\"dxa\" w:w=\"#{total_width / columns}\"", @noko))
				body_cell_borders = body_cell_props.add_child(Nokogiri::XML::Node.new("tcBorders", @noko))
				body_cell_borders.add_child(Nokogiri::XML::Node.new("top w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				body_cell_borders.add_child(Nokogiri::XML::Node.new("left w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				body_cell_borders.add_child(Nokogiri::XML::Node.new("bottom w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				body_cell_borders.add_child(Nokogiri::XML::Node.new("right w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				body_cell_props.add_child(Nokogiri::XML::Node.new("shd w:fill=\"000000\" w:val=\"clear\"", @noko))
				body_cell_margins = body_cell_props.add_child(Nokogiri::XML::Node.new("tcMar", @noko))
				body_cell_margins.add_child(Nokogiri::XML::Node.new("left w:type=\"dxa\" w:w=\"55\"", @noko))
				body_paragraph = body_cell.add_child(Nokogiri::XML::Node.new("p", @noko))
				body_paragraph_props = body_paragraph.add_child(Nokogiri::XML::Node.new("pPr", @noko))
				body_paragraph_props.add_child(Nokogiri::XML::Node.new("pStyle w:val=\"style20\"", @noko))
				body_paragraph_props.add_child(Nokogiri::XML::Node.new("rPr", @noko))
				body_paragraph_run = body_paragraph.add_child(Nokogiri::XML::Node.new("r", @noko))
				body_paragraph_run.add_child(Nokogiri::XML::Node.new("rPr", @noko))
			end
		end
		
		@state.document = @noko.to_s()
		puts "Document: #{@state.document}\n\n"
	end
	
=begin
 <w:tr>
	<w:trPr>
		<w:cantSplit w:val="false"/>
	</w:trPr>
	<w:tc>
		<w:tcPr>
			<w:tcW w:type="dxa" w:w="3212"/>
			<w:tcBorders>
				<w:top w:color="000000" w:space="0" w:sz="2" w:val="single"/>
				<w:left w:color="000000" w:space="0" w:sz="2" w:val="single"/>
				<w:bottom w:color="000000" w:space="0" w:sz="2" w:val="single"/>
				<w:right w:val="nil"/>
			</w:tcBorders>
			<w:shd w:fill="FFFF00" w:val="clear"/>
			<w:tcMar>
				<w:left w:type="dxa" w:w="54"/>
			</w:tcMar>
		</w:tcPr>
		<w:p>
			<w:pPr>
				<w:pStyle w:val="style20"/>
				<w:rPr/>
			</w:pPr>
			<w:r>
				<w:rPr/>
			</w:r>
		</w:p>
	</w:tc>
	<w:tc>
		<w:tcPr>
			<w:tcW w:type="dxa" w:w="3213"/>
			<w:tcBorders>
				<w:top w:color="000000" w:space="0" w:sz="2" w:val="single"/>
				<w:left w:color="000000" w:space="0" w:sz="2" w:val="single"/>
				<w:bottom w:color="000000" w:space="0" w:sz="2" w:val="single"/>
				<w:right w:val="nil"/>
			</w:tcBorders>
			<w:shd w:fill="FFFF00" w:val="clear"/>
			<w:tcMar>
				<w:left w:type="dxa" w:w="54"/>
			</w:tcMar>
		</w:tcPr>
		<w:p>
			<w:pPr>
				<w:pStyle w:val="style20"/>
				<w:rPr/>
			</w:pPr>
			<w:r>
				<w:rPr/>
			</w:r>
		</w:p>
	</w:tc>
	<w:tc>
		<w:tcPr>
			<w:tcW w:type="dxa" w:w="3213"/>
			<w:tcBorders>
				<w:top w:color="000000" w:space="0" w:sz="2" w:val="single"/>
				<w:left w:color="000000" w:space="0" w:sz="2" w:val="single"/>
				<w:bottom w:color="000000" w:space="0" w:sz="2" w:val="single"/>
				<w:right w:color="000000" w:space="0" w:sz="2" w:val="single"/>
			</w:tcBorders>
			<w:shd w:fill="FFFF00" w:val="clear"/>
			<w:tcMar>
				<w:left w:type="dxa" w:w="54"/>
			</w:tcMar>
		</w:tcPr>
		<w:p>
			<w:pPr>
				<w:pStyle w:val="style20"/>
				<w:rPr/>
			</w:pPr>
			<w:r>
				<w:rPr/>
				<w:t>Yellow row</w:t>
			</w:r>
		</w:p>
	</w:tc>
</w:tr>
=end
	

end


















