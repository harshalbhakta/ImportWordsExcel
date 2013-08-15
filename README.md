Import words from Excel
==

This is a simple ruby program to generate INSERT queries from a excel sheet.

Excel Details
-------------------------

The [excel sheet](https://github.com/harshalbhakta/ImportWordsExcel/blob/master/wtw_words_demo.xls?raw=true) has below mentioned columns.

* A:1  - Word
* A:2  - Pronunciation
* A:3  - Sentence1
* A:4  - Sentence2
* A:5  - Sentence3
* A:6  - Sentence4
* A:7  - Sentence5
* A:8  - Sentence6
* A:9  - Sentence7
* A:10 - Sentence8
* A:11 - Sentence9
* A:12 - Sentence10
* A:13 - Sentence11
* A:14 - Sentence12
* A:15 - Sentence13
* A:16 - Sentenct14

Parsing Rules
-------------------------

* Each word has one pronunciation & multiple sentences associated with it.
* We need atleast 5 sentences to accept a word entry.
* We stop parsing when we get a empty word entry.
* We skip sentences that are empty.

Queries to be generated
-------------------------

*Word query*

```sql

INSERT INTO words (id,word,pronunciation,ishidden,created_at,updated_at) VALUES 
                  ('1','a la carte','ah-luh-KAHRT','false','2013-08-14 00:00:00','2013-08-14 00:00:00');

```
