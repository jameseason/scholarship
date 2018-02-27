# Topic Frequency Analyzer 

Use user-generated topical categories to run an analysis on abstract data. Generates an excel file tallying the frequency of articles with one or more hits in a given topic in a given year. 

## Topics
A topic consists of a topic title and any number of items. Items may be (1) whole words, (2) phrases (“rural sociology”), or (3) parts of words with an asterisk * (“farm*” will match “farmer”, “farming”, etc). These can be imported or exported as a text file.


## Analysis
The analysis tallies the frequency of articles with one or more hits in a given topic in a given year. The output is an excel file (xlsx) containing the number of appearances of each topic in each year, with all found years as the first column, and all specified topics as the first row.

## Assumptions
All abstracts will be a text file (txt) with a title starting with its year of publication. Topics files will be text files following the format defined in sample_topics.txt.
