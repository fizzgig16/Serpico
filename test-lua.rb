#!/usr/bin/env ruby

require 'zipruby'
require './helpers/lua_parser'

document = ""
Zip::Archive.open("templates/testdoc.docx", Zip::CREATE) do |zipfile|
		# read in document.xml and store as var
		zipfile.fopen("word/document.xml") do |f|
			document = f.read # read entry content
		end
	end

noko = Nokogiri::XML(document)
run_lua(noko, "")
