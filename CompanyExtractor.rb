require 'csv'
require 'pdf-reader-turtletext'
require 'spreadsheet'

## pagenum_filename - csv list of pages to read from pdf file. (Created in CompanyPageFinder.rb)
## pdf_filename     - Combined pdf file of all the separate company pdf files.
## company_pages    - Stores the page numbers to read from in the pdf file.
## book        - A new excel workbook
## sheet       - A new sheet in the excel workbook
## current_row - Keeps track of which row to write to in the excel sheet.
pagenum_filename = File.join(File.expand_path File.dirname(__FILE__), 'pdf_output2')
pdf_filename = File.join(File.expand_path File.dirname(__FILE__), 'subset.pdf')
company_pages = []
book = Spreadsheet::Workbook.new # Create a new excel file to write to.
sheet = book.create_worksheet
current_row = 0


# Reading in the PDF page numbers that contain company information into an array.
CSV.foreach(pagenum_filename, "r") do |row|
	company_pages.push(row)
end

## reader - used to read the pdf file
reader = PDF::Reader::Turtletext.new(pdf_filename) # Takes a long time becuase of pdf size

company_pages.each do |page|

	# Extracts the company_name form a specific region on the page.
	company_name = reader.text_in_region(0, 900, 535, 800, page[0].to_i)[0][0]

	## region - The specific pixel region on the page that has the company information.
	region = reader.bounding_box do
		page    page[0].to_i
		below   "Headquarters Contact Information"
		above   320
		left_of 400
	end

	## text_in_region - An array of text lines extracted from the region.
	text_in_region = region.text

	## current_index - Keeps track of the current position while looping through text_region.
	## phone_index   - Holds the element index of the phone number.
	current_index = 0
	phone_index = 0

	# Loops through the text_in_region array to find the element index of the phone number.
	text_in_region.each do |text|
		if text.to_s.include? "Phone"
			phone_index = current_index
		else
			current_index += 1
		end
	end

	## phone - contains the company phone number.
	## address - Contains the company address.
	## employees - Contains the amount of people employed by the company.
	## website - Contains the company's website.

	# Removes the text from phone number line. (Example- Phone: +1-402-324-3435 -> 1-402-324-3435)
	phone = text_in_region[phone_index][0].gsub(/[^-0-9]/, '')

	# Concatenates all elements of the text_in_region array before the phone_index to form the full address.
	address = ""
	0.upto(phone_index - 1) { |i| address += " "; address += text_in_region[i][0].to_s }

	# Removes the text from the employees line. (Example- Employees: 2200 -> 2200)
	employees = text_in_region[phone_index+2][0].gsub(/[^0-9]/, '').to_i
	
	website = text_in_region[phone_index+1][0]

	# Writes the company information to a row in the excel sheet.
	sheet.row(current_row).push company_name, address, phone, website, employees

	puts current_row  # displays progress of the script on the command line.
	current_row += 1  # iterates to next row in the excel sheet.
end

# save file
book.write File.join(File.expand_path File.dirname(__FILE__), 'results.xls')

