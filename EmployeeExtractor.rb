require 'csv'
require 'pdf-reader-turtletext'
require 'spreadsheet'
require 'debugger'

## pagenum_filename - csv list of pages to read from pdf file. (Created in CompanyPageFinder.rb)
## pdf_filename     - Combined pdf file of all the separate company pdf files.
## employee_pages   - Stores the page numbers to read from in the pdf file.
## book        - A new excel workbook
## sheet       - A new sheet in the excel workbook
## current_row - Keeps track of which row to write to in the excel sheet.
pagenum_filename = File.join(File.expand_path File.dirname(__FILE__), 'EmployeePages')
pdf_filename = File.join(File.expand_path File.dirname(__FILE__), 'subset.pdf')
employee_pages = []
book = Spreadsheet::Workbook.new # Create a new excel file to write to.
sheet = book.create_worksheet
current_row = 0
contacts = []
regex = Regexp.new('(\d*[A-Z][a-z]+(\s\d*[A-Z][a-z]+)*)')
phone_regex = Regexp.new('(\d{1}[\s\d-]+)')
email_regex = Regexp.new('/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b/i')


# Reading in the PDF page numbers that contain employee names into an array.
CSV.foreach(pagenum_filename, "r") do |row|
	employee_pages.push(row)
end

## reader - used to read the pdf file
reader = PDF::Reader::Turtletext.new(pdf_filename) # Takes a long time because of the pdf's size

# Iterate through each page that contacts employee information.
employee_pages.each do |page|


	## region - The specific pixel region on the page that has the employee names.
	region = reader.bounding_box do
		page    page[0].to_i
		below   /Company/
		#above   320
		#left_of 200
		inclusive true
	end

	debugger

	# Loops through each row (which is one contact) to extract the contact's info.
	region.text.each do |row|
		contact = []
		debugger
		row.each do |piece|
			scanner = piece.scan(regex)
			scanner.each do |string|
				contact.push(string.first) if string.first != nil
			end
		end	
			
		contact.pop # remove last because it isn't a contact

		phone = row.last.scan(phone_regex).flatten.first.to_s

		email = row.last.scan(email_regex)
		email = email.first.to_s
		email_start = email.index(/[A-Za-z]/).to_i
		email = email[email_start, email.length]

		contact.push(phone)
		contact.push(email)

		contacts.push(contact)
	end 
end


contacts.each do |row|
	begin
		if (row.length > 5) && (row.first != "Company") && (row.first != "Corporate Overview")

			company = row.first
			row.shift
			
			name = row[0] + " " + row[1]
			row.shift
			row.shift
			
			email = row.last
			row.pop
			
			phone = row.last
			row.pop

			rank = row.last
			row.pop
			
			title = row.join(", ")		

			# Writes the company information to a row in the excel sheet.
			sheet.row(current_row).push company, name, title, rank, phone, email

			puts current_row  # displays progress of the script on the command line.
			current_row += 1  # iterates to next row in the excel sheet.
		end
	rescue
		puts "1"
	end
end

# save file
book.write File.join(File.expand_path File.dirname(__FILE__), 'theresults.xls')