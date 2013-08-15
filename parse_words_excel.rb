require 'spreadsheet'

book = Spreadsheet.open('wtw_words_demo.xls')
words_sheet = book.worksheet('words') 		# can use an index or worksheet name

content = ""
processed_words = 0
sequence_no = 0
words_sheet.each_with_index do |row,index|

  next if index == 0						# Skip the header row
  
  word = row[0] ? row[0].strip : nil  
  break if (word.nil? || word.strip == "")	# Stop parsing if word field is empty.

  pronunciation = row[1] ? row[1].strip : ""

  sentences = []
  (2..15).each do |row_index|				# Total 14 options

  	sentence = row[row_index] ? row[row_index].strip : ""

  	# Skip if minimum 10 sentences are not available.
  	break if (row_index < 7 && sentence == "")

  	# Skip empty sentences after 5th sentence.
  	next if (sentence == "")

  	sentences.push(sentence)
  end

  # Skip the word if there are less than 5 options.
  if sentences.count < 5
  	puts "Skipping '#{word}' as it has only #{sentences.count} options."
  	next
  end

  # Generate INSERT query for word
  content += %Q{INSERT INTO words (id,word,pronunciation,ishidden,created_at,updated_at) VALUES ('#{index}','#{word.gsub("'","''")}','#{pronunciation.gsub("'","''")}','false','2013-08-14 00:00:00','2013-08-14 00:00:00');} + "\n\n"

  # Gemerate INSERT query for sentences
  sentences.each do |sentence|
  	content += %Q{INSERT INTO sentences (sentence,word_id,created_at,updated_at) VALUES ('#{sentence.gsub("'","''")}','#{index}','2013-08-14 00:00:00','2013-08-14 00:00:00');} + "\n"
  	sequence_no = index + 1
  end

  processed_words += 1
  content += "\n"

end

puts "Insert queries generated for #{processed_words} words."

content += ("\n\n\n" + "SELECT setval('words_id_seq',#{sequence_no}, true);")

file = File.open("output.txt", "w")
file.write(content)