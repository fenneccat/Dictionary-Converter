# Dictionary-Converter

This program makes MS Words vocabulary file with words that you wrote on excel file.

Run program by
```console
python Dictionary_converter.py
```

You need to install `python-docx` library

If you put a name of book or movie after you run the program, it will automatically create the folder and files.
While watching video or reading book, if you found some unknown vocabulary, write it on `bookname_new.xlsx`.
Foreign word should located on first column and translation should locate on second column.

Whenever you write and save the `bookname_new.xlsx` file, run the program to update your vocabulary file.

In bookname_Dictionary MS Words file, words are sorted in confusing order (frequency of wrting certain word) in first section.
Some words are turn red when newly saved word is already registered on your Dictionary already.

In second section, all words that you wrote is sorted in alphabet order.

Enjoy your language learning!
