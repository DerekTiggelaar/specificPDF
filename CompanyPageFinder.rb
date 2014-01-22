require 'csv'
require 'pdf-reader'

pagenum_filename = File.join(File.expand_path File.dirname(__FILE__), 'pdf_output2')
pdf_filename = File.join(File.expand_path File.dirname(__FILE__), 'subset.pdf')
@reader = PDF::Reader.new(pdf_filename)
@writer = CSV.open(pagenum_filename, 'wb')
@mutex = Mutex.new
threads = []
@current_page = 1
@total_pages = 500 # Length of pdf file


def find_companies
	while @current_page != @total_pages
		read_page
	end
end

def read_page
	@mutex.synchronize do
		page = @reader.page(@current_page)
		if page.text.include? "Company Profile"
			if !page.text.include? "TABLE OF CONTENTS" 
				@writer << [@current_page]
			end
		end
		puts @current_page
		@current_page += 1
	end
end

1.upto(20) do |thread|
	threads << Thread.new { Thread.current['id'] = thread; find_companies }
end
threads.each { |t| t.join }

@writer.close
@reader.close