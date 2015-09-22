# encoding: ASCII-8BIT

# Responsible for parsing Lua content in Serpico templates

require 'rubygems'
require 'nokogiri'
require 'rlua'
require 'cgi'
require 'json'
require './model/master.rb'

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
		
	doc_text = CGI::unescapeHTML(document.to_s()).force_encoding("ASCII-8BIT")
	
	# Build a list of block locations with a «. We will use those as insertion points for where the code lives, and thus where content should be placed
	lua_block_loc_nodes = []
	lua.get_nokogiri_document().xpath("//w:t[contains(text(), \"«\")]", 'w' => $W_URI).each do |lua_block_node|
		#puts "Adding #{lua_block_node.to_s()}"
		lua_block_loc_nodes << lua_block_node
	end
	
	block_loc = 0
	doc_text.scan(/«(.*?)»/m).each do |lua_block|
		lua_block = lua_block[0]	# Only get the capture group

		# Replace paragraph ends with newlines. This is a trick for our regex below
		lua_block.gsub!(/<\/w:p>/, "<w:t>\\n</w:t>")
		
		# Grab <w:t> blocks and piece them together
		code = ""
		lua_block.scan(/<w:t[^\w]?.*?>(.*?)<\/w:t>/).each do |text_block|
			code = code + text_block[0].to_s().gsub(/\\n/, "\n")
		end
		
		code = code.force_encoding("ASCII-8BIT")
		#puts "Found Lua block: #{code}"
		lua.run_lua_block(code, lua_block_loc_nodes[block_loc])
		block_loc = block_loc + 1
	end
	
	# All Lua blocks are done, get rid of any Lua stuff
	lua.get_nokogiri_document().xpath("//w:t[contains(text(), \"«\")]", 'w' => $W_URI).each do |lua_start_node|
		# Nuke any siblings up to and including anything that contains »
		#puts "found start #{lua_start_node.to_s()}"
		
		found_end = false		
		node = new_node = lua_start_node
		
		until found_end do
			while new_node do
				#puts "nuking node #{new_node.to_s()}"
				if new_node.content.force_encoding("ASCII-8BIT").include?("»")
					# End of our content, delete and stop the loop
					#puts "found end"
					found_end = true
					new_node.content = ""
					break
				end
				
				new_node.content = ""
				new_node = new_node.next_sibling
			end
			
			unless found_end
				# Still haven't found the end, move up to the parent paragraph and try again
				foo = node.parent.next_sibling
				#puts "foo is #{foo.to_s()}"
				unless foo
					# If it's nil, nothing else in this run, move up one more
					#puts "moving to gparent"
					foo = node.parent.parent.next_sibling
				end
				#puts "going to parent #{foo.to_s()}"
				node = new_node = foo.xpath(".//w:t", 'w' => $W_URI).first
				unless node
					# No text nodes here, go to next sibling
					node = new_node = foo.next_sibling
				end
				#puts "new node: #{new_node.to_s()}"
			end
		end
	end
	
	return lua.get_nokogiri_document()
end

class SerpicoLua
	def initialize(doc, report_xml)
		# Start the Lua engine
		@state = Lua::State.new()
		
		# Build a nokogiri document from the doc we passed in so we only parse it once
		@noko = doc
		
		# Store report XML
		report_xml = report_xml
		
		# Build the nokogiri doc for the report XML too
		@noko_report = Nokogiri::XML(report_xml)
		
		# Now set up our tables so that we can call stuff
		create_lua_tables()
	end

	def run_lua_block(lua_block, block_loc)
		#puts "block_loc: #{block_loc.to_s()}"
		# Clean up stupid stuff like smart quotes and double-dashes
		clean_block = lua_block
		clean_block.gsub!("“", "\"")
		clean_block.gsub!("”", "\"")
		clean_block.gsub!("–", "--")
		
		# Run ze code
		@loc_to_insert = block_loc
		@state.__eval(clean_block)
		
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
			# ReportContent:GetReportVars()
			'GetReportVars' => lambda { |this| lua_get_report_vars(this) }
		}
		@state.Findings = 
		{
			# Findings:GetAllFindings() - returns a table of Findings
			'GetAllFindings' => lambda { |this| lua_findings_getallfindings(this) }
		}
			
		@state.Finding = 
		{
			# Finding:GetTitle(finding)
			'GetTitle' => lambda { |this| lua_finding_gettitle(this) },
			'GetID' => lambda { |this| lua_finding_getid(this) }
		}

		@state.WordTable =
		{
			# WordTable:CreateSimple(columns, header_row_count, body_row_count)
			'CreateSimple' => lambda { |this, columns, header_row_count, body_row_count| lua_wordtable_create_simple(this, columns, header_row_count, body_row_count) },
		
			# WordTable:Create(columns, header_row_count, body_row_count)
			'Create' => lambda { |this, columns, header_row_count, body_row_count, border_width, border_color, cell_border_width, cell_border_color, some_hash| lua_wordtable_create(this, columns, header_row_count, 
		                          body_row_count, border_width, border_color, cell_border_width, cell_border_color, some_hash) }
		}
			
		@state.Document =
		{
			# WordTable:Create(columns, header_row_count, body_row_count)
			'AddParagraph' => lambda { |this, text| lua_document_addparagraph(this, text) }
		}
			
		@state.User = 
		{
			# User:GetDetails(username)
			'GetDetails' => lambda { |this, username| lua_user_info(username) },
			
			# User:GetAllContributors()
			'GetAllContributors' => lambda { |this| lua_user_contributors() }
		}
		
	end

	# Gets the label complete with delimiters
	def get_full_label(label)
		return "¤#{label}¤"
	end

	#### Label functions ####
	# Replaces an entire label with another value. This destroys the label, so be careful when you call it!
	def lua_label_replace(this, labelname, value)
		#puts "Called Label:Replace"
		full_label = get_full_label(labelname)
		@noko.xpath('//text()').each do |node|
			node.content = node.content.force_encoding("ASCII-8BIT").gsub(/#{full_label}/, value)
			#puts "Replacing content with #{value}"
		end
	end

	# Sets the foreground color of the label text
	def lua_label_forecolor(this, labelname, rgbstring)
		full_label = get_full_label(labelname)
		
		# Need to see if our label is immediately preceded by <w:t>; if so, means we have our own text block we can colorize, or at least star with
		# If not, need to end the block we're currently in and start a new one with the color of our choosing
		@noko.xpath("//w:t[contains(text(), \"#{full_label}\")]", 'w' => $W_URI).each do |text_node|
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
		end
	end
	
	#### ReportContent functions ####
	def lua_reportcontent_getshortcompanyname(this)
		return @noko_report.xpath("/report/reports/short_company_name").first.content
	end
	
	#### Findings functions ####
	
	# Returns a table of Finding(s)
	def lua_findings_getallfindings(this)
		output = []
		findings = @noko_report.xpath("/report/findings_list/findings")
		findings.each do |finding|
			# Create a table
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
			#table["risk"] = finding.xpath("risk").first.content
			
			output << table
		end
		
		return output
	end
	
	def lua_get_report_vars(this)
		table = {}
		
		var = @noko_report.xpath("/report/reports")
		if var
			# Create a table

			# Functions
			table["GetDate"] = lambda { |thistable| thistable["date"] }
			table["GetReportType"] = lambda { |thistable| thistable["report_type"] }
			table["GetReportName"] = lambda { |thistable| thistable["report_name"] }
			table["GetConsultantName"] = lambda { |thistable| thistable["consultant_name"] }
			table["GetConsultantPhone"] = lambda { |thistable| thistable["consultant_phone"] }
			table["GetConsultantTitle"] = lambda { |thistable| thistable["consultant_title"] }
			table["GetConsultantEmail"] = lambda { |thistable| thistable["consultant_email"] }
			table["GetContactName"] = lambda { |thistable| thistable["contact_name"] }
			table["GetContactPhone"] = lambda { |thistable| thistable["contact_phone"] }
			table["GetContactTitle"] = lambda { |thistable| thistable["contact_title"] }
			table["GetContactEmail"] = lambda { |thistable| thistable["contact_email"] }
			table["GetContactCity"] = lambda { |thistable| thistable["contact_city"] }
			table["GetContactAddress"] = lambda { |thistable| thistable["contact_address"] }
			table["GetContactState"] = lambda { |thistable| thistable["contact_state"] }
			table["GetContactZip"] = lambda { |thistable| thistable["contact_zip"] }
			table["GetFullCompanyName"] = lambda { |thistable| thistable["full_company_name"] }
			table["GetShortCompanyName"] = lambda { |thistable| thistable["short_company_name"] }
			table["GetCompanyWebsite"] = lambda { |thistable| thistable["company_website"] }
			
			# Attributes
			table["date"] = CGI::unescapeHTML(var.xpath("date").first.content)
			table["report_type"] = CGI::unescapeHTML(var.xpath("report_type").first.content)
			table["report_name"] = CGI::unescapeHTML(var.xpath("report_name").first.content)
			table["consultant_name"] = CGI::unescapeHTML(var.xpath("consultant_name").first.content)
			table["consultant_phone"] = CGI::unescapeHTML(var.xpath("consultant_phone").first.content)
			table["consultant_title"] = CGI::unescapeHTML(var.xpath("consultant_title").first.content)
			table["consultant_email"] = CGI::unescapeHTML(var.xpath("consultant_email").first.content)
			table["contact_name"] = CGI::unescapeHTML(var.xpath("contact_name").first.content)
			table["contact_phone"] = CGI::unescapeHTML(var.xpath("contact_phone").first.content)
			table["contact_title"] = CGI::unescapeHTML(var.xpath("contact_title").first.content)
			table["contact_email"] = CGI::unescapeHTML(var.xpath("contact_email").first.content)
			table["contact_city"] = CGI::unescapeHTML(var.xpath("contact_city").first.content)
			table["contact_address"] = CGI::unescapeHTML(var.xpath("contact_address").first.content)
			table["contact_state"] = CGI::unescapeHTML(var.xpath("contact_state").first.content)
			table["contact_zip"] = CGI::unescapeHTML(var.xpath("contact_zip").first.content)
			table["full_company_name"] = CGI::unescapeHTML(var.xpath("full_company_name").first.content)
			table["short_company_name"] = CGI::unescapeHTML(var.xpath("short_company_name").first.content)
			table["company_website"] = CGI::unescapeHTML(var.xpath("company_website").first.content)
			if (var.xpath("user_defined_variables") != nil)
				# User vars are in JSON format. Add each one to the lua table
				user_vars = JSON.parse(var.xpath("user_defined_variables").first.content)
				user_vars.each do |key,value|
					#puts "Found user var #{key} with value #{CGI::unescapeHTML(value)}"
					table["#{key}"] = CGI::unescapeHTML(value)
				end
			end
			
			# Todo: add authors and user defined vars
			#table["owner"] = var.xpath("owner").first.content
			#table["contact_state"] = var.xpath("contact_state").first.content
		end
		
		return table
	end
	
	#### WordTable functions ####
	def lua_wordtable_create(this, columns, header_row_count, body_row_count, border_width, border_color, cell_border_width, cell_border_color, some_hash)
		# Convert strings to ints
		columns = columns.to_i()
		header_row_count = header_row_count.to_i()
		body_row_count = body_row_count.to_i()
		border_width = border_width.to_i()
		cell_border_width = cell_border_width.to_i()
		arrWidths = []
		
		# Creates a table in word where the code is
		total_width = 4500
		table_id = rand(999999).to_s()

		#puts "Creating table with ID #{table_id}"
		if some_hash != nil
			# Unpack the column_widths table, which are nested, thus the temp "each" loop
			if some_hash["column_widths"] != nil
				some_hash["column_widths"].each do |dont_care,temp|
					temp.each do |still_dont_care, width|
						#puts("Width: #{width}")
						arrWidths << width.to_i()
					end
				end
				if arrWidths.length != columns
					puts "column_widths and column count are different, ignoring widths parameter (#{columns}/#{arrWidths.length})"
					arrWidths = nil
				end
			end
		end
		
		# This is the table we will pass back to Lua for future reference
		result = {}
		result["id"] = table_id
		
		#table = @noko.xpath("//w:body").first.add_child(Nokogiri::XML::Node.new("w:tbl", @noko))
		table = @loc_to_insert.add_previous_sibling(Nokogiri::XML::Node.new("w:tbl", @noko))
		
		# Add custom property to the table so we can find it later
		table["w:rsidR"] = table_id # WARNING: this is technically not valid, as the DTD does not define such an attribute, but it seems to not crash anything. It also means we can't search it directly
		
		table_params = table.add_child(Nokogiri::XML::Node.new("w:tblPr", @noko))
		table_params.add_child(Nokogiri::XML::Node.new("tblW w:type=\"pct\" w:w=\"#{total_width}\"", @noko))
		table_params.add_child(Nokogiri::XML::Node.new("jc w:val=\"left\"", @noko))
		table_params.add_child(Nokogiri::XML::Node.new("tblInd w:type=\"dxa\" w:w=\"55\"", @noko))
		if (border_width > 0)
			table_borders = table_params.add_child(Nokogiri::XML::Node.new("tblBorders", @noko))
			table_borders.add_child(Nokogiri::XML::Node.new("top w:color=\"#{border_color}\" w:space=\"0\" w:sz=\"#{border_width}\" w:val=\"single\"", @noko))
			table_borders.add_child(Nokogiri::XML::Node.new("left w:color=\"#{border_color}\" w:space=\"0\" w:sz=\"#{border_width}\" w:val=\"single\"", @noko))
			table_borders.add_child(Nokogiri::XML::Node.new("bottom w:color=\"#{border_color}\" w:space=\"0\" w:sz=\"#{border_width}\" w:val=\"single\"", @noko))
			table_borders.add_child(Nokogiri::XML::Node.new("insideH w:color=\"#{border_color}\" w:space=\"0\" w:sz=\"#{border_width}\" w:val=\"single\"", @noko))
			table_borders.add_child(Nokogiri::XML::Node.new("right w:val=\"nil\"", @noko))
			table_borders.add_child(Nokogiri::XML::Node.new("insideV w:val=\"nil\"", @noko))
		end
		cell_margins = table_params.add_child(Nokogiri::XML::Node.new("tblCellMar", @noko))
		cell_margins.add_child(Nokogiri::XML::Node.new("top w:type=\"dxa\" w:w=\"55\"", @noko))
		cell_margins.add_child(Nokogiri::XML::Node.new("left w:type=\"dxa\" w:w=\"55\"", @noko))
		cell_margins.add_child(Nokogiri::XML::Node.new("bottom w:type=\"dxa\" w:w=\"55\"", @noko))
		cell_margins.add_child(Nokogiri::XML::Node.new("right w:type=\"dxa\" w:w=\"55\"", @noko))
		table_grid = table.add_child(Nokogiri::XML::Node.new("tblGrid", @noko))

		for cols in 1..columns
			width = (total_width / columns).floor	# This is the default
			if arrWidths.length > 0
				width = arrWidths[cols - 1].to_i()
				#puts "Setting width[#{cols - 1}] to #{width}"
			end
			table_grid.add_child(Nokogiri::XML::Node.new("gridCol w:w=\"#{width}\"", @noko))
		end
	
		# Now the rows, header first
		result["header_rows"] = []
		#puts "Header count: #{header_row_count}"
		#puts "Body count: #{body_row_count}"
		
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
				width = (total_width / columns).floor	# This is the default
				if arrWidths.length > 0
					width = arrWidths[cols - 1].to_i()
					#puts "Setting header width[#{cols - 1}] to #{width}"
				end
				header_cell = header_row.add_child(Nokogiri::XML::Node.new("tc", @noko))
				header_cell_props = header_cell.add_child(Nokogiri::XML::Node.new("tcPr", @noko))
				header_cell_props.add_child(Nokogiri::XML::Node.new("tcW w:type=\"pct\" w:w=\"#{width}\"", @noko))
				if (cell_border_width > 0)
					header_cell_borders = header_cell_props.add_child(Nokogiri::XML::Node.new("tcBorders", @noko))
					header_cell_borders.add_child(Nokogiri::XML::Node.new("top w:color=\"#{cell_border_color}\" w:space=\"0\" w:sz=\"#{cell_border_width}\" w:val=\"single\"", @noko))
					header_cell_borders.add_child(Nokogiri::XML::Node.new("left w:color=\"#{cell_border_color}\" w:space=\"0\" w:sz=\"#{cell_border_width}\" w:val=\"single\"", @noko))
					header_cell_borders.add_child(Nokogiri::XML::Node.new("bottom w:color=\"#{cell_border_color}\" w:space=\"0\" w:sz=\"#{cell_border_width}\" w:val=\"single\"", @noko))
					header_cell_borders.add_child(Nokogiri::XML::Node.new("right w:color=\"#{cell_border_color}\" w:space=\"0\" w:sz=\"#{cell_border_width}\" w:val=\"single\"", @noko))
				end
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
				width = (total_width / columns).floor	# This is the default
				if arrWidths.length > 0
					width = arrWidths[cols - 1].to_i()
					#puts "Setting body width[#{cols - 1}] to #{width}"
				end
				body_cell = body_row.add_child(Nokogiri::XML::Node.new("tc", @noko))
				body_cell_props = body_cell.add_child(Nokogiri::XML::Node.new("tcPr", @noko))
				body_cell_props.add_child(Nokogiri::XML::Node.new("tcW w:type=\"pct\" w:w=\"#{width}\"", @noko))
				if (cell_border_width > 0)
					body_cell_borders = body_cell_props.add_child(Nokogiri::XML::Node.new("tcBorders", @noko))
					body_cell_borders.add_child(Nokogiri::XML::Node.new("top w:color=\"#{cell_border_color}\" w:space=\"0\" w:sz=\"#{cell_border_width}\" w:val=\"single\"", @noko))
					body_cell_borders.add_child(Nokogiri::XML::Node.new("left w:color=\"#{cell_border_color}\" w:space=\"0\" w:sz=\"#{cell_border_width}\" w:val=\"single\"", @noko))
					body_cell_borders.add_child(Nokogiri::XML::Node.new("bottom w:color=\"#{cell_border_color}\" w:space=\"0\" w:sz=\"#{cell_border_width}\" w:val=\"single\"", @noko))
					body_cell_borders.add_child(Nokogiri::XML::Node.new("right w:color=\"#{cell_border_color}\" w:space=\"0\" w:sz=\"#{cell_border_width}\" w:val=\"single\"", @noko))
				end
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
	
	def lua_wordtable_create_simple(this, columns, header_row_count, body_row_count)
		return lua_wordtable_create_simple(this, columns, header_row_count, body_row_count, 2, "000000", 2, "000000")		
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
		
		# Now find the table
		table_root = nil	
		@noko.xpath("//w:tbl", 'w' => $W_URI).each do |attr_node|
			if (attr_node.to_s().include?("w:rsidR=\"#{table["id"]}\""))
				table_root = attr_node
				break
			end
		end

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
		
		row = @noko.xpath("//w:tr[@w:rsidTr=\"#{row["id"]}\"]/w:tc[#{cellindex}]", 'w' => $W_URI).first
		if row
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
			
			last_br_node = nil
			text.split("\n").each do |line|
				if line
					if text_node == nil
						text_node = run.add_child(Nokogiri::XML::Node.new("t", @noko))
					end
					text_node.content = line.force_encoding("ASCII-8BIT")
					last_br_node = text_node.add_next_sibling(Nokogiri::XML::Node.new("br", @noko))
					text_node = nil
				end
			end
			
			if last_br_node != nil
				last_br_node.remove()	# Get rid of the empty line at the end
			end
			
			return 1
		end
		
		return nil
	end
	
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
		
		cell_node = @noko.xpath("//w:tr[@w:rsidTr=\"#{row["id"]}\"]/w:tc[#{cellstartindex}]/w:tcPr", 'w' => $W_URI).first
		if cell_node
			cell_node.add_child(Nokogiri::XML::Node.new("gridSpan w:val=\"#{numcols}\"", @noko))
			
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
		#para_prop.add_child(Nokogiri::XML::Node.new("pStyle", @noko))
		para_prop.add_child(Nokogiri::XML::Node.new("rPr", @noko))
		run = para.add_child(Nokogiri::XML::Node.new("r", @noko))
		run.add_child(Nokogiri::XML::Node.new("rPr", @noko))
		text_node = run.add_child(Nokogiri::XML::Node.new("t", @noko))
		text_node.content = text.force_encoding("ASCII-8BIT")
		
		return 1
	end
	
	def lua_user_info(username)
		#puts "Selecting user #{username}"
		
		user = User.first(:username => username)
		if user == nil
			puts "No users found"
			return 0
		end
		
		lua_user = {}
		lua_user["id"] = user.id
		lua_user["consultant_name"] = user.consultant_name
		lua_user["consultant_email"] = user.consultant_email
		lua_user["consultant_title"] = user.consultant_title
		lua_user["consultant_phone"] = user.consultant_phone
		#puts "Found a user, name is #{user.consultant_name}"
		
		return lua_user
	end
	
	def lua_user_contributors()
		all_users = User.all
		authors = @noko_report.xpath("//authors").first.content
		if authors
			authors = authors.gsub(/[\[\]]/,'')
			authors = authors.gsub(/\"/,'')
			auth_array = authors.split(",")
			lua_contributors = []
			auth_array.each do |author|
				lua_contributors << author.strip()
			end
			
			return lua_contributors
		end
		
		return 0
	end
end


















