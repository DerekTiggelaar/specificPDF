require 'csv'
require 'pdf-reader'
require 'spreadsheet'

## pagenum_filename - csv list of pages to read from pdf file. (Created in CompanyPageFinder.rb)
## pdf_filename     - Combined pdf file of all the separate company pdf files.
## employee_pages   - Stores the page numbers to read from in the pdf file.
## book        - A new excel workbook
## sheet       - A new sheet in the excel workbook
## current_row - Keeps track of which row to write to in the excel sheet.
@dir = File.expand_path File.dirname(__FILE__)
pagenum_filename = File.join(@dir, 'EmployeePagesSplit.csv')
pdf_filename = File.join(@dir, 'combined_pdfs.pdf')
employee_pages = []
book = Spreadsheet::Workbook.new # Create a new excel file to write to.
sheet = book.create_worksheet
current_row = 0
contacts = []
regex = Regexp.new('([A-Z][a-z]+(\s\d*[A-Z][a-z]+)*)')
phone_regex = Regexp.new('(\d{1}[\s\d-]+)')
email_regex = Regexp.new('/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b/i')
company_regex = Regexp.new('([A-Z][a-z]+)|([A-Z]+)')

class Contact
	attr_accessor :company, :name, :email, :phone, :rank, :title
end

# Reading in the PDF page numbers that contain employee names into an array.
CSV.foreach(pagenum_filename, "r") do |row|
	employee_pages.push(row)
end

## reader - used to read the pdf file
#reader = PDF::Reader::Turtletext.new(pdf_filename) # Takes a long time because of the pdf's size
reader = PDF::Reader.new(pdf_filename)

# Iterate through each page that contacts employee information.
employee_pages.each do |page|

	current_page = reader.page(page[0].to_i).text
	text = current_page.each_line.to_a
	company_name = current_page.split("\n")[1]

	# Loops through each row (which is one contact) to extract the contact's info.
	text.each do |row|
		contact = Contact.new
		info = []

		begin

			# Scans the row that looks for camelCases. 
			# The PDF doesn't separate the rows, we can tell different rows b/c no space separates a capitalized word from the end previous.
			scanner = row.scan(regex)
			scanner.each do |string|
				info.push(string.first) if string.first != nil
			end


			if (info.length > 3) && (scanner.first.first != "Company") && (scanner.first.first != "Corporate Overview")

				contact.company = company_name

				company_name_split = company_name.scan(company_regex).join(" ")


				#company_name_split = company_name.gsub("-", " ")
				#company_name_split = company_name_split.delete("^\sa-zA-Z0-9")
				company_name_split = company_name_split.split

				company_name_split.each do |sub_string|
					if info.first.include? sub_string
						info.shift
					end
				end

				first_name = info.shift
				last_name = info.shift
				contact.name = first_name + " " + last_name
				

				contact.phone = row.scan(phone_regex).flatten.first.to_s

				email = row.split(" ").last
				email_start = email.index(/[A-Za-z]/).to_i
				if email_start != 0
					contact.email = email[email_start, email.length]
				end

				info.pop
				contact.rank = info.pop.to_s
				contact.title = info.join(", ")
				contacts.push(contact)	
			end
		rescue Exception => err
			puts err
		end
	end 
end

contacts.each do |contact|
	# Writes the company information to a row in the excel sheet.
	sheet.row(current_row).push contact.company, contact.name, contact.title, contact.rank, contact.phone, contact.email
	current_row += 1  # iterates to next row in the excel sheet.
end

puts current_row.to_s + " full contacts found." # displays progress of the script on the command line.

# save file
book.write File.join(@dir, 'resultsPartTwo.xls')