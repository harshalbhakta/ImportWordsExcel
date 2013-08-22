Import words from Excel
==

This is a simple ruby program to generate INSERT queries from a excel sheet.

Excel Details
-------------------------

The [excel sheet](https://github.com/harshalbhakta/ImportWordsExcel/blob/master/wtw_words_demo.xls?raw=true) has below mentioned columns.

| Publish | Word  | Procunciation | Sentence1  | Sentence2 | ... | Sentence14 |
| ------- | ----- | :-----------: | :--------: | :-------: | :-: | ---------: |
|    Y	  | abase | /əˈbās/       | Sample text| Sample text| ... | Last Sentence|

* A:1  - publish (y or n)
* A:2  - Word
* A:3  - Pronunciation
* A:4  - Sentence1
* A:5  - Sentence2
* A:6  - Sentence3
* A:7  - Sentence4
* A:8  - Sentence5
* A:9  - Sentence6
* A:10 - Sentence7
* A:11 - Sentence8
* A:12 - Sentence9
* A:13 - Sentence10
* A:14 - Sentence11
* A:15 - Sentence12
* A:16 - Sentence13
* A:17 - Sentenct14

Parsing Rules
-------------------------

* Each word has one pronunciation & multiple sentences associated with it.
* We stop parsing when we get a empty word entry.
* We skip sentences that are empty.

Queries to be generated
-------------------------

Word query

```sql

INSERT INTO words (id,word,pronunciation,approved,published,created_at,updated_at) VALUES ('1','a la carte','ah-luh-KAHRT','true','true','2013-08-22 21:20:18','2013-08-22 21:20:18');

```

Sentence query

```sql

INSERT INTO sentences (sentence,word_id,created_at,updated_at) VALUES ('Utilize these resources to complete the application and/or make enhancements to   the school''s menu or a la carte  program.','1','2013-08-22 21:20:18','2013-08-22 21:20:18');

INSERT INTO sentences (sentence,word_id,created_at,updated_at) VALUES ('The menu has both combination plates and a la carte  options.','1','2013-08-22 21:20:18','2013-08-22 21:20:18');

```

Reset word sequence query

```sql

# words_id_seq is last word count + 1
SELECT setval('words_id_seq',20, true);

```

