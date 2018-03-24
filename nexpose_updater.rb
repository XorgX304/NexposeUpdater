#!/usr/bin/env ruby
require 'rubygems'
require 'Nexpose'
require 'roo'
require 'io/console'

# Variables for Nexpose Instance
host = "10.1.106.8"
port = "3780"
user = ""
pass = ""

# Start Main...

# Connect and Authenticate
begin
	# Ask for Username & Password to avoid storing clear text credentials in script
	print "Please Enter your Username: "
	user = STDIN.noecho(&:gets).chomp
	puts ""
	print "Please Enter your Password: "
	pass = STDIN.noecho(&:gets).chomp
	puts ""

	# Create a connection to the Nexpose instance
	@nsc = Nexpose::Connection.new(host, user, pass, port)

	# Authenticate to this instance (throws an exception if this fails)
	@nsc.login
	
rescue ::Nexpose::APIError => e
	$stderr.puts ("Connection failed: #{e.reason}")
	exit(1)
end

# Open Unidentified Assets XLSX file
print "Please Enter the XLSX File or Path Name of your Unidentified Assets List: "
xlsx_file = STDIN.gets.chomp
puts
1.times do
	file_exist = File.exist?(xlsx_file)
	if file_exist
		@xlsx = Roo::Spreadsheet::open(xlsx_file)
	else
		sleep 1
		puts "This file doesn't exist!"
		sleep 1
		puts
		print "Please Enter the XLSX File or Path Name of your Unidentified Assets List: "
		xlsx_file = STDIN.gets.chomp
		puts
		redo
	end
end

# Query a list of all Assets in Nexpose from the XLSX file
begin
@xlsx.each do |row|
	if row.at(1) == "Host IP" or row.at(7) == "Asset Owner" 
		# Don't Print Headers
	else
        ip = row.at(1)
        host_name = row.at(2)
        owner = row.at(7)
		
		# Filter for Asset in Nexpose via IP
		assets = @nsc.filter(Nexpose::Search::Field::IP_RANGE, Nexpose::Search::Operator::IS, [ip, ip])

		if assets.empty?
			puts "#{ip} not found!"
		else
			assets.each do |asset|
				puts "#{asset.ip} - Found!"
				
        # TODO - FIX LINES 75-97
				# Check current tagged Owners in Asset and adds them to an array 
				#set_tags = @nsc.list_asset_tags(asset.id)
				#preset_tags = []
				#set_tags.each do |t|
				#	preset_tags.push(t.name)
				#end
				
				# Check current tagged Owners in XLSX, splits them, and adds them to an array
				#xlsx_tags = []
				#unless owner.nil?
				#	owner.split(",").each do |t|
				#		xlsx_tags.push(t)
				#	end
				#end
				
				# If Owners in XLSX != Owners in Asset, remove differences 
				#remove_owner = preset_tags - xlsx_tags
				#remove_owner.each do |o|
				#	tag_id = @nsc.list_tags.find { |t| t.name == o }.id rescue nil # Catch Nil Error
				#	@nsc.remove_tag_from_asset(asset.id, tag_id)
				#	puts "#{o} is no longer an Owner for #{asset.ip} and has been removed"
				#end
				
				# Continue to add owners from XLSX
				if owner == nil
					# If Owner Column is Empty - skip tagging
					puts "No Owner Tag Provided!"
				else
					# Split string if there is more than 1 Owner
					owner.split(",").each do |o|
						owner = o
						# If Owner column contains a Name then check for the Owner Tag ID in Nexpose
						tag_id = @nsc.list_tags.find { |t| t.name == owner }.id rescue nil # Catch Nil Error
					
						# If Owner Tag ID does not exist (Nil Error) then create a new Owner Tag in Nexpose
						if tag_id.nil?
							puts "No Owner for #{ip}"
							puts "Creating a new Owner Tag..."
							new_tag = Nexpose::Tag.new(owner, "OWNER")
							id = new_tag.save(@nsc)
							puts "New Tag '#{owner}' saved with ID: #{id}"
						
							#Add New Owner Tag to Asset
							tag = Nexpose::Tag.load(@nsc, id)
							tag.add_to_asset(@nsc, asset.id)
							puts "New Tag '#{tag.name}' added to #{asset.ip}"
						else
						# If Owner Tag ID does exist then load the ID Tag and add it to the Asset
							unless @nsc.asset_tags(asset.id).find { |t| t.id == tag_id }
								tag = Nexpose::Tag.load(@nsc, tag_id)
								tag.add_to_asset(@nsc, asset.id)
								puts "New Tag '#{tag.name}' added to #{asset.ip}"
							end
						end
					end
				end
			end
		end
	end
	puts "--------------------------------------------------------------"
end
end
