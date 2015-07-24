# encoding: ASCII-8BIT

# Responsible for parsing Lua content in Serpico templates

require 'rubygems'
require 'nokogiri'
require 'rlua'

$W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

class SerpicoLua
	def initialize(doc)
		# Start the Lua engine
		@state = Lua::State.new()
		
		# Store document in state so we can work with it
		@state.document = doc
		
		# Build a nokogiri document from the doc we passed in so we only parse it once
		@noko = Nokogiri::XML(@state.document)
		
		# Now set up our tables so that we can call stuff
		create_lua_tables()
	end

	def run_lua_block(lua_block)
		# Clean up stupid stuff like smart quotes and double-dashes
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
		@state.__load_stdlib :math, :string, :table	# Only load bare minimum tables
		#@state.__eval "print('hello, world')" # launch some Lua code
		
		# Create tables we can call from lua to do stuff
		create_label_table()
	end

	def create_label_table()
		@state.Label = 
		{
			# Label:Replace(labelname, value)
			'Replace' => lambda { |this, labelname, value| lua_label_replace(this, labelname, value) },
			# Label:ForeColor(labelname, rgbstring) - RGB string will be something like 0070c0
			'ForeColor' => lambda { |this, labelname, rgbstring| lua_label_forecolor(this, labelname, rgbstring) }
		}
	end

	# Gets the label complete with delimiters
	def get_full_label(label)
		return "¤#{label}¤"
	end

	# Replaces an entire label with another value. This destroys the label, so be careful when you call it!
	def lua_label_replace(this, labelname, value)
		full_label = get_full_label(labelname)
		@state.document = @state.document.gsub!(/#{full_label}/, value)
	end

	# Sets the foreground color of the label text
	def lua_label_forecolor(this, labelname, rgbstring)
		full_label = get_full_label(labelname)
		document = @state.document
		
		# Need to see if our label is immediately preceded by <w:t>; if so, means we have our own text block we can colorize, or at least star with
		# If not, need to end the block we're currently in and start a new one with the color of our choosing
		#puts "Noko: #{@noko}"
		@noko.xpath("//w:t[contains(text(), \"#{full_label}\")]", 'w' => $W_URI).each do |text_node|
			#puts "Found #{text_block}"
			run_props = text_node.xpath("../w:rPr", 'w' => $W_URI).first
			run_props.add_child(Nokogiri::XML::Node.new("color w:value=\"#{rgbstring}\"", @noko))
			parent = text_node.xpath("..", 'w' => $W_URI)
			puts "My parent is #{parent}"
			if (text_node.content == full_label)
				# Text is by itself, we can apply the color directly
				
			else
				
			end
			
			#puts "My parent is #{parent}"
		end
		
		@state.document = @state.document.gsub!(/#{full_label}/, "<w:color w:val=\"#{rgbstring}\"/>#{full_label}")
	end
			
end