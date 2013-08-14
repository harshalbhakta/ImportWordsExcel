require 'spreadsheet'    

book = Spreadsheet.open('wtw_words.xls')

words_sheet = book.worksheet('words') # can use an index or worksheet name

content = ""
words_sheet.each do |row|
  content += ( row.join(',') + "\n" )
end

file = File.open("output.txt", "w")
file.write(content)