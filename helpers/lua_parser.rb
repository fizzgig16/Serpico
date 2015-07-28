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
			'GetAllFindings' => lambda { |this| Lua.multret(lua_findings_getallfindings(this)) }
		}
			
		@state.Finding = 
		{
			# Finding:GetTitle(finding)
			'GetTitle' => lambda { |this| return lua_finding_gettitle(this) },
			'GetID' => lambda { |this| return lua_finding_getid(this) }
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
				prop_node.add_child(Nokogiri::XML::Node.new("shd w:fill=\"#{rgbstring}\" v:val=\"clear\"", @noko))
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
					prop_node.add_child(Nokogiri::XML::Node.new("shd w:fill=\"#{rgbstring}\" v:val=\"clear\"", @noko))
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
		puts "GetAllFindings()"
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
	
end


















