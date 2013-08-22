require 'spreadsheet'
require 'active_support/time'

### Check Arguments
if ARGV[0] == nil
  puts "Please provide the filename of the xls sheet."
  exit
end

file_name = ARGV[0]
words_sheet = Spreadsheet.open(file_name).worksheet(0) 		# can use an index or worksheet name
content = ""                                              # will construct the string for output file.
processed_words = 0
today_string = DateTime.now.change(:offset => "+0530").strftime('%Y-%m-%d %T')

# Column indexes
publish_col        = 0
word_col           = 1
pronunciation_col  = 2
sentence_start_col = 3
sentence_end_col   = 16

def text_for_cell(cell)
  cell ? cell.strip : ""
end

# Iterate through rows
words_sheet.each_with_index do |row,index|

  next if index == 0						# Skip the header row
  
  word = text_for_cell(row[word_col])
  pronunciation = text_for_cell(row[pronunciation_col])
  approved = text_for_cell(row[publish_col]).downcase == "y" ? "true" : "false"
  
  break if (word == "")	        # Stop parsing if word field is empty.

  sentences = []
  (sentence_start_col..sentence_end_col).each do |row_index|
  	sentence = text_for_cell(row[row_index])
  	next if (sentence == "") # Skip empty sentences
  	sentences.push(sentence)
  end

  # Generate INSERT query for word
  content += %Q{INSERT INTO words (id,word,pronunciation,approved,published,created_at,updated_at) VALUES ('#{index}','#{word.gsub("'","''")}','#{pronunciation.gsub("'","''")}','#{approved}','true','#{today_string}','#{today_string}');} + "\n\n"
  
  # Gemerate INSERT query for sentences
  sentences.each do |sentence|
  	content += %Q{INSERT INTO sentences (sentence,word_id,created_at,updated_at) VALUES ('#{sentence.gsub("'","''")}','#{index}','#{today_string}','#{today_string}');} + "\n"
  end

  processed_words += 1
  content += "\n"

end

puts "Insert queries generated for #{processed_words} words."

processed_words += 1

content += ("\n\n\n" + "SELECT setval('words_id_seq',#{processed_words}, true);")

file = File.open(file_name.gsub(".xls","_output.txt"), "w")
file.write(content)