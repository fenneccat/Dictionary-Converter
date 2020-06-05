# Dictionary-Converter

This program makes MS Words vocabulary file with words that you wrote on excel file.

## Quick Start
Run program by
```console
python Dictionary_converter.py
```

## Requirement
You need to install `python-docx` library

## Fucntion Explanation
If you put a name of book or movie after you run the program, it will automatically create the folder and files. 

When you first put a book/movie name, it will ask what language do you want to learn. Source Language (Learning Language) is a language of contents (book, movie, etc) and Target Language (Native Language) is a language of that you knows well. 

While watching a video or reading a book, if you found some unknown vocabulary, write it on `bookname_new.xlsx`.
Foreign word should located on first column and translation in native language should locate on second column.

Whenever you write and save the `bookname_new.xlsx` file, run the program to update your vocabulary file.

* Input Example

![image of input](https://github.com/fenneccat/Dictionary-Converter/blob/master/images/new_data.JPG)

DataBase will be automatically generated. You don't have to care about it. DB looks like bellow
* Database example

![image of DB](https://github.com/fenneccat/Dictionary-Converter/blob/master/images/DB.JPG)

In first section, newly added words are listed with blue color.

In bookname_Dictionary MS Words file, words are sorted in confusing order (frequency of wrting certain word) in second section.
Some words are turn red when newly saved word is already registered on your Dictionary already.

In third section, all words that you wrote is sorted in alphabet order.

* Generated Vocabulary File

![image of word](https://github.com/fenneccat/Dictionary-Converter/blob/master/images/screen.JPG)

Enjoy your language learning!
