require 'csv'
require 'pdf-reader'

@dir = File.expand_path File.dirname(__FILE__)
pagenumber_file = File.join(@dir, 'theSubsetPages')
pdf_filename = File.join(@dir, 'subset.pdf')
@reader = PDF::Reader.new(pdf_filename)
@mutex = Mutex.new
threads = []
@current_page = 1
@total_pages = 500 # Length of pdf file

class Page
	attr_accessor :number, :content
	
	def initialize(number, content)
		@number = number
		@content = content
	end
end

def find_employees
	valid = CSV.open(File.join(@dir, "valid#{Thread.current['id']}.tmp"), 'wb')
	begin
		while @current_page < @total_pages
			page = read_page
			if page.content.include? "ORG CHARTS AND CONTACTS"
				if !page.content.include? "full org charts" 
					valid << [page.number]
				end
			end
		end
	rescue Exception => err
		puts err
	end
	valid.close
end

def read_page
	page = nil
	@mutex.synchronize do
		if @current_page < @total_pages
			page_content = @reader.page(@current_page).text
			page = Page.new(@current_page, page_content)
			puts @current_page
			@current_page += 1
		end
	end
	page
end

def combine_files(pattern, src_file)
	temp_files = Dir.glob( File.join(@dir, pattern) )
	begin 
		open(src_file, 'w') do |file|
			temp_files.each do |f|
				file.write(IO.read(f))
			end
		end
	rescue Exception => err
		puts "Error combining temp files - #{err.message}"
	end
end

1.upto(10) do |thread|
	threads << Thread.new { Thread.current['id'] = thread; find_employees }
end
threads.each { |t| t.join }

combine_files('*.tmp', pagenumber_file)

puts 'done'