=====================================
WARNING: VERY ALPHA SOFTWARE - DO NOT USE FOR SENSITIVE TASKS


This software is a very basic tool for making and manipulating indexes of folders on, currently, Windows computers.

It is really two tools in one:

1) a tool to create an index of a given folder.  The index is outputted to an Excel file, and can optionally include additional information from certain file formats (currently just emails)
2) a tool to rename files based on an Excel index (including one you might generate under 1).  Users can set proposed file names, then run the program to rename the files en masse.

The use of the term 'rename' is a bit of a stretch - for safety, all the files will be copied to a new folder with '_renamed' appended to it.  Maybe in a future release I will be brave enough to provide a 'rename in place' option, but it is not this day.

The program is a series of basic python files, wrapped with a wxPython GUI. It is not the 

Just to be clear: this is not the most intelligent software out there. It does not rely on an LLM to read a document, intuit its contents and come up with a file name. Instead, it provides the user with a platform to be able to easily and flexibly traverse large sets of documents and, if it would be helpful, rename those files.

===================================================
Why did you make this?
===================================================

The purpose for the software was to solve three very specific problems:

1) first, being able to very quickly create indexes of directories, particularly for the purpose of creating workflows for reviewing documents
2) second, being able to quickly rename files en masse, especially after being able to leverage Excel's formulas and other tools to quickly create standardised file name types
3) third, and related to 2), being able to automatically extract information from certain files to enable files to be quickly renamed - for example, intelligently naming emails based on their senders and recipients.

You might think this is not a common problem, but believe me - in certain industries (especially the law) this is _very_ common.  Seeing one too many junior lawyers creating an index table from scratch for a discovery exercise is what finally triggered me to write this code.

===================================================
Did you use AI to assist in creating the code
===================================================

Yes.  There are many aspects of programming process that are frustrating, and which LLMs are particularly well suited for.  For example:

1) The program uses wxPython which, while universal, requires a certain amount of arcane magic in order to set up a basic window interface. The examples, while useful, are also time-consuming to implement. Having an LLM being able to generate the template GUI code was extremely helpful.
2) Quickly being able to come up with simple 'helper' functions, such as being able to remove problematic characters from file name strings. These are very common pieces of code that LLMs can pluck out of their data sets very easily, but which finding out in the wild may be quite difficult.
3) Finally, programming is not my job.  I am reasonably adept in Python, and while I have proven to myself I can certainly work out how to write the code from scratch without LLM assistance, in some circumstances it is too slow. In this case, I needed this code for some paid work, and as much as I would have been intellectually satisfied with the outcome had I coded it entirely from scratch, in practice I needed the code to 'get out of the way' so that I could get on with my value add.

===================================================
What is there left to do
===================================================

Gosh, so much.

The current 0.1 roadmap is as follows:

1) Create more versatile controls for suggesting proposed filenames for emails (currently there is one format hard-coded into 'list-documents.py')
2) Cross-platform support (which should notionally be easy given the use of the standard wxPython and pathlib libraries)
3) Make the output index files more readable
4) Make the file name editing process more resilient (e.g. by making it less easy to omit file extensions and bork your files)
5) Make the different 'panels' interact with each other easier - if you create an index in one session, it should remember the file and folder locations for the renamer tool
6) Make the program in general more resilient - and include some logs or error messages that don't fail silently
7) Make the interface prettier
8) And no doubt more...

===================================================
License
===================================================

Please use this code if it is helpful for you - my only request is that it doesn't make its way into some closed-source project that people charge for. It is also not good enough to make it into some commercial code - if you would like to sell a product like this, please go and make something better from scratch.

Many times, I have come across situations in which computers are the enemy - for example, struggling with Windows refusing to rename a file, or being forced to do some turgid, awful task that I could have sworn could be automated, but can't be because a tool costs money (which the business won't pay for).  I strongly believe that computers and code should liberate people, not imprison them. 

Please note the license file for the relevant restrictions.
