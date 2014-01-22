require 'csv'
require 'pdf-reader'
require 'debugger'

pagenumber_file = File.join(File.expand_path File.dirname(__FILE__), 'theSubsetPages')
pdf_filename = File.join(File.expand_path File.dirname(__FILE__), 'subset.pdf')
@reader = PDF::Reader.new(pdf_filename)
@writer = CSV.open(pagenumber_file, 'wb')
@mutex = Mutex.new
threads = []
@current_page = 1
@total_pages = 500 # Length of pdf file


def find_employees
	while @current_page != @total_pages
		page = read_page
		if page.nil? != true
			page_content = @reader.page(page)
			if page_content.text.include? "ORG CHARTS AND CONTACTS"
				if !page_content.text.include? "full org charts" 
					@writer << [page]
				end
			end
		end
	end
end

def read_page
	page = nil
	@mutex.synchronize do
		begin
			page = @current_page
		rescue
			puts "Page " + @current_page.to_s + " did not work."
		end
		puts @current_page
		@current_page += 1
	end
	page
end

1.upto(20) do |thread|
	threads << Thread.new { Thread.current['id'] = thread; find_employees }
end
threads.each { |t| t.join }

@writer.close
@reader.close