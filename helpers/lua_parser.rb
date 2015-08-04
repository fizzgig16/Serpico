# encoding: ASCII-8BIT

# Responsible for parsing Lua content in Serpico templates

require 'rubygems'
require 'nokogiri'
require 'rlua'
require 'cgi'

$W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
# o thing: &#xA4;
# Start lua: &#xAB;
# End Lua: &#xBB;
def run_lua(document, report_xml)
	###############################
	# « - begin Lua block
	# » - end Lua block
	
	# Find any labels we might need later
	#build_label_array(document)
	
	lua = SerpicoLua.new(document, report_xml)
		
	#puts "Doc before: #{document}\n\n"
	#puts "Report: #{report_xml}\n\n"
	doc_text = CGI::unescapeHTML(document.to_s()).force_encoding("ASCII-8BIT")
	#puts doc_text
	doc_text.scan(/«(.*?)»/m).each do |lua_block|
		lua_block = lua_block[0]	# Only get the capture group
		
		#puts "Before: #{lua_block}"
		
		# Replace paragraph ends with newlines. This is a trick for our regex below
		lua_block.gsub!(/<\/w:p>/, "<w:t>\\n</w:t>")
		
		# Grab <w:t> blocks and piece them together
		code = ""
		lua_block.scan(/<w:t[^\w]?.*?>(.*?)<\/w:t>/).each do |text_block|
			code = code + text_block[0].to_s().gsub(/\\n/, "\n")
		end
		
		code = code.force_encoding("ASCII-8BIT")
		#puts "Found Lua block: #{code}"
		lua.run_lua_block(code)
	end
	
	# After our processing is done, replace document with whatever is in the Lua state
	doc_text = lua.get_document()
	#puts document
	
	return lua.get_nokogiri_document()
end

class SerpicoLua
	def initialize(doc, report_xml)
		# Start the Lua engine
		@state = Lua::State.new()
		
		# Store document in state so we can work with it
		#@state.document = doc
		
		# Build a nokogiri document from the doc we passed in so we only parse it once
		@noko = doc
		
		# Store report XML
		report_xml = report_xml
		
		# Build the nokogiri doc for the report XML too
		@noko_report = Nokogiri::XML(report_xml)
		
		# Now set up our tables so that we can call stuff
		create_lua_tables()
	end

	def run_lua_block(lua_block)
		# Clean up stupid stuff like smart quotes and double-dashes
		
		#puts "Calling Lua block"
		
		clean_block = lua_block
		clean_block.gsub!("“", "\"")
		clean_block.gsub!("”", "\"")
		clean_block.gsub!("–", "--")
		
		# Run ze code
		@state.__eval(clean_block)
		
	end

	def get_document()
		return ""
	end
	
	def get_nokogiri_document()
		return @noko
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
			
		@state.Document =
		{
			# WordTable:Create(columns, header_row_count, body_row_count)
			'AddParagraph' => lambda { |this, text| lua_document_addparagraph(this, text) }
		}
		
	end

	# Gets the label complete with delimiters
	def get_full_label(label)
		return "¤#{label}¤"
	end

	#### Label functions ####
	# Replaces an entire label with another value. This destroys the label, so be careful when you call it!
	def lua_label_replace(this, labelname, value)
		#puts "Called replace"
		full_label = get_full_label(labelname)
		@noko.xpath('//text()').each do |node|
			node.content = node.content.force_encoding("ASCII-8BIT").gsub(/#{full_label}/, value)
		end
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
		end
	end
			
	#### Cell functions ####
	def lua_cell_backcolor(this, labelname, rgbstring)
		full_label = get_full_label(labelname)
		
		@noko.xpath("//w:tc/w:p/w:r/w:t[contains(text(), \"#{full_label}\")]", 'w' => $W_URI).each do |cell_node|
			#puts "Found #{cell_node}"
			shading_node = cell_node.xpath("../../../w:tcPr/w:shd", 'w' => $W_URI).first
			if (shading_node)
				#puts "Found shading node"
				#shading_node = shading_node.first
				shading_node["w:fill"] = rgbstring
			else
				prop_node = cell_node.xpath("../../../w:tcPr", 'w' => $W_URI).first
				#<w:shd w:fill="FF0000" w:val="clear"/>
				prop_node.add_child(Nokogiri::XML::Node.new("shd w:fill=\"#{rgbstring}\" w:val=\"clear\"", @noko))
			end
			
			#puts "Content: #{@noko}\n\n"
			#@state.document = @noko.to_s()
		end
	end
	
	#### Row functions ####
	def lua_row_backcolor(this, labelname, rgbstring)
		full_label = get_full_label(labelname)
		
		@noko.xpath("//w:tr/w:tc/w:p/w:r/w:t[contains(text(), \"#{full_label}\")]", 'w' => $W_URI).each do |row_node|
			#puts "Found #{row_node}"
			cell_nodes = row_node.xpath("../../../../w:tc", 'w' => $W_URI)
			cell_nodes.each do |cell_node|
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
			#@state.document = @noko.to_s()
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
			table["GetTitle"] = lambda { |table| table["title"] }
			table["GetID"] = lambda { |table| table["id"] }
			table["GetEffort"] = lambda { |table| table["effort"] }
			table["GetType"] = lambda { |table| table["type"] }
			table["GetOverview"] = lambda { |table| table["overview"] }
			table["GetRemediation"] = lambda { |table| table["remediation"] }
			table["GetRisk"] = lambda { |table| table["risk"] }
			
			# Attributes
			table["id"] = finding.xpath("id").first.content
			table["title"] = finding.xpath("title").first.content
			table["effort"] = finding.xpath("effort").first.content
			table["type"] = finding.xpath("type").first.content
			table["overview"] = finding.xpath("overview/paragraph").first.content
			table["remediation"] = finding.xpath("remediation/paragraph").first.content
			table["risk"] = finding.xpath("risk").first.content
			
			output << table
		end
		
		return output
	end
	
	#### WordTable functions ####
	def lua_wordtable_create(this, columns, header_row_count, body_row_count)

		# Convert strings to ints
		columns = columns.to_i()
		header_row_count = header_row_count.to_i()
		body_row_count = body_row_count.to_i()

		# Creates a table in word where the code is
		total_width = 9681
		table_id = rand(999999).to_s()

		puts "Creating table with ID #{table_id}"
		
		# This is the table we will pass back to Lua for future reference
		result = {}
		result["id"] = table_id
		
		table = @noko.xpath("//w:body").first.add_child(Nokogiri::XML::Node.new("w:tbl", @noko))
		
		# Add custom property to the table so we can find it later
		table["w:rsidR"] = table_id # WARNING: this is technically not valid, as the DTD does not define such an attribute, but it seems to not crash anything. It also means we can't search it directly
		
		table_params = table.add_child(Nokogiri::XML::Node.new("w:tblPr", @noko))
		table_params.add_child(Nokogiri::XML::Node.new("tblW w:type=\"dxa\" w:w=\"#{total_width}\"", @noko))
		table_params.add_child(Nokogiri::XML::Node.new("jc w:val=\"left\"", @noko))
		table_params.add_child(Nokogiri::XML::Node.new("tblInd w:type=\"dxa\" w:w=\"55\"", @noko))
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
			table_grid.add_child(Nokogiri::XML::Node.new("gridCol w:w=\"#{(total_width / columns).floor - 18}\"", @noko))
		end
	
		# Now the rows, header first
		result["header_rows"] = []
		#puts "Have #{header_row_count} headers"
		
		for header_rows in 1..header_row_count
			# Add the header table to the results
			header = {}
			header["id"] = rand(999999).to_s()
			header["index"] = header_rows
			header["SetCellText"] = lambda { |this, cellindex, text| lua_row_setcelltext(this, cellindex, text) }
			header["SpanCells"] = lambda { | this, cellstartindex, numcols | lua_row_spancells(this, cellstartindex, numcols) }
			result["header_rows"] << header
			
			header_row = table.add_child(Nokogiri::XML::Node.new("tr", @noko))
			
			# Add custom property to the table so we can find it later
			header_row["w:rsidTr"] = header["id"]
			
			header_props = header_row.add_child(Nokogiri::XML::Node.new("trPr", @noko))
			header_props.add_child(Nokogiri::XML::Node.new("tblHeader w:val=\"true\"", @noko))
			header_props.add_child(Nokogiri::XML::Node.new("cantSplit w:val=\"false\"", @noko))
			
			# Add the cells
			for cols in 1..columns
				header_cell = header_row.add_child(Nokogiri::XML::Node.new("tc", @noko))
				header_cell_props = header_cell.add_child(Nokogiri::XML::Node.new("tcPr", @noko))
				header_cell_props.add_child(Nokogiri::XML::Node.new("tcW w:type=\"dxa\" w:w=\"#{(total_width / columns).floor - 18}\"", @noko))
				header_cell_borders = header_cell_props.add_child(Nokogiri::XML::Node.new("tcBorders", @noko))
				header_cell_borders.add_child(Nokogiri::XML::Node.new("top w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				header_cell_borders.add_child(Nokogiri::XML::Node.new("left w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				header_cell_borders.add_child(Nokogiri::XML::Node.new("bottom w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				header_cell_borders.add_child(Nokogiri::XML::Node.new("right w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				header_cell_props.add_child(Nokogiri::XML::Node.new("shd w:fill=\"auto\" w:val=\"clear\"", @noko))
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
		result["body_rows"] = []
		for body_rows in 1..body_row_count
			# Add the body table to the results
			body = {}
			body["id"] = rand(999999).to_s()
			body["index"] = body_rows
			body["SetCellText"] = lambda { |this, cellindex, text| lua_row_setcelltext(this, cellindex, text) }
			body["SpanCells"] = lambda { | this, cellstartindex, numcols | lua_row_spancells(this, cellstartindex, numcols) }
			result["body_rows"] << body
			
			body_row = table.add_child(Nokogiri::XML::Node.new("tr", @noko))
			
			# Add custom property to the table so we can find it later
			body_row["w:rsidTr"] = body["id"]
			
			body_props = body_row.add_child(Nokogiri::XML::Node.new("trPr", @noko))
			body_props.add_child(Nokogiri::XML::Node.new("cantSplit w:val=\"false\"", @noko))
			
			# Add the cells
			for cols in 1..columns
				body_cell = body_row.add_child(Nokogiri::XML::Node.new("tc", @noko))
				body_cell_props = body_cell.add_child(Nokogiri::XML::Node.new("tcPr", @noko))
				body_cell_props.add_child(Nokogiri::XML::Node.new("tcW w:type=\"dxa\" w:w=\"#{(total_width / columns).floor - 18}\"", @noko))
				body_cell_borders = body_cell_props.add_child(Nokogiri::XML::Node.new("tcBorders", @noko))
				body_cell_borders.add_child(Nokogiri::XML::Node.new("top w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				body_cell_borders.add_child(Nokogiri::XML::Node.new("left w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				body_cell_borders.add_child(Nokogiri::XML::Node.new("bottom w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				body_cell_borders.add_child(Nokogiri::XML::Node.new("right w:color=\"000000\" w:space=\"0\" w:sz=\"2\" w:val=\"single\"", @noko))
				body_cell_props.add_child(Nokogiri::XML::Node.new("shd w:fill=\"auto\" w:val=\"clear\"", @noko))
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
		
		# Functions
		result["GetBodyRow"] = lambda { |table, index| lua_wordtable_getbodyrow(table, index) }
		result["AddBodyRow"] = lambda { |table| lua_wordtable_addbodyrow(table) }
			
		return result
	end
	
	def lua_wordtable_getbodyrow(table, index)
		index = index.to_i()
		
		# Start by making sure the table ID is a number
		unless table["id"].to_i() > 0
			# throw an exception or something?
			return nil
		end
		
		unless index > 0
			# throw an exception or something?
			return nil
		end
		
		#puts "Looking for table #{table["id"]}"
		
		# Now find the table
		table_root = nil	
		@noko.xpath("//w:tbl", 'w' => $W_URI).each do |attr_node|
			if (attr_node.to_s().include?("w:rsidR=\"#{table["id"]}\""))
				table_root = attr_node
				break
			end
		end

		#puts table_root
		unless table_root
			# throw an exception or something?
			return nil
		end
		
		# Got it
		row = table_root.xpath("w:tr[#{index}]", 'w' => $W_URI).first
		if row
			body = {}
			body["id"] = row["w:rsidTr"]
			body["index"] = index
			body["SetCellText"] = lambda { |this, cellindex, text| lua_row_setcelltext(this, cellindex, text) }
			body["SpanCells"] = lambda { | this, cellstartindex, numcols | lua_row_spancells(this, cellstartindex, numcols) }
			
			return body
		end
		
		return nil
	end
	
	def lua_wordtable_addbodyrow(table)
		# Start by making sure the table ID is a number
		unless table["id"].to_i() > 0
			# throw an exception or something?
			return nil
		end
		
		# Now find the table
		table_root = nil	
		@noko.xpath("//w:tbl", 'w' => $W_URI).each do |attr_node|
			if (attr_node.to_s().include?("w:rsidR=\"#{table["id"]}\""))
				table_root = attr_node
				break
			end
		end

		#puts table_root
		unless table_root
			# throw an exception or something?
			return nil
		end
		
		# Copy the last <w:tr> and append to the table
		old_row = table_root.xpath("w:tr[last()]", 'w' => $W_URI).first
		if old_row
			new_row = old_row.clone(1)
			
			# Replace the ID with a new one
			new_id = rand(999999).to_s()
			#puts "New row should be ID #{new_id}"
			new_row["w:rsidTr"] = new_id
			table_root.add_child(new_row)
		else
			# throw an exception or something?
			return nil
		end

		return 1
	end

	def lua_row_setcelltext(this, cellindex, text)
		row = this
		
		unless row["id"].to_i() > 0
			# throw an exception or something?
			return nil
		end
		
		# Got it
		#puts "valid row"
		
		row = @noko.xpath("//w:tr[@w:rsidTr=\"#{row["id"]}\"]/w:tc[#{cellindex}]", 'w' => $W_URI).first
		if row
			#puts "Found row"
			
			# Create w:p if it doesn't exist
			para = row.xpath("w:p", 'w' => $W_URI).first
			if para == nil
				para = row.add_child(Nokogiri::XML::Node.new("p", @noko))
			end
			
			# Create w:r if it doesn't exist
			run = para.xpath("w:r", 'w' => $W_URI).first
			if run == nil
				run = para.add_child(Nokogiri::XML::Node.new("r", @noko))
			end
			
			# Create w:t if it doesn't exist
			text_node = run.xpath("w:t", 'w' => $W_URI).first
			if text_node == nil
				text_node = run.add_child(Nokogiri::XML::Node.new("t", @noko))
			end
			
			if text
				text_node.content = text.force_encoding("ASCII-8BIT")
			end
			
			return 1
		end
		
		return nil
	end
	
	#<w:gridSpan w:val="2"/>
	def lua_row_spancells(this, cellstartindex, numcols)
		row = this
		
		unless row["id"].to_i() > 0
			# throw an exception or something?
			return nil
		end
		
		cellstartindex = cellstartindex.to_i()
		numcols = numcols.to_i()
		unless cellstartindex > 0
			return nil
		end
		unless numcols > 0
			return nil
		end
		
		# Got it
		puts "valid row"
		
		cell_node = @noko.xpath("//w:tr[@w:rsidTr=\"#{row["id"]}\"]/w:tc[#{cellstartindex}]/w:tcPr", 'w' => $W_URI).first
		if cell_node
			puts "found cell: #{cell_node.to_s()}\n\n"
			cell_node.add_child(Nokogiri::XML::Node.new("gridSpan w:val=\"#{numcols}\"", @noko))
			puts "New node: #{cell_node.to_s()}\n\n"
			
			# Now get rid of numcols - 1 w:tc nodes after it
			for i in 1..(numcols -1)
				cell_next = cell_node.parent.next_sibling.remove()
			end
		end
		
		return 1
	end
	
	def lua_document_addparagraph(this, text)
		body = @noko.xpath("//w:document/w:body", 'w' => $W_URI).first
		unless body
			# throw an exception or something?
			return nil
		end
		
		para = body.add_child(Nokogiri::XML::Node.new("p", @noko))
		para_prop = para.add_child(Nokogiri::XML::Node.new("pPr", @noko))
		para_prop.add_child(Nokogiri::XML::Node.new("pStyle", @noko))
		para_prop.add_child(Nokogiri::XML::Node.new("rPr", @noko))
		run = para.add_child(Nokogiri::XML::Node.new("r", @noko))
		run.add_child(Nokogiri::XML::Node.new("rPr", @noko))
		text_node = run.add_child(Nokogiri::XML::Node.new("t", @noko))
		text_node.content = text.force_encoding("ASCII-8BIT")
		
		#puts @noko.to_s()
		
		return 1
	end
	
end


















