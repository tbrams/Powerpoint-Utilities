#PowerPoint Utility
This is where I will add utilities that I use a lot in Powerpoint and at some point ... maybe sometime in the future, this will become a nicely 
integrated tool

Two things can be accomplised here:
- Automatic inserting and scaling pictures 
- Securing powerpoint comments

##Automatic inserting pictures
For now this is the startup form - can be changed easily... but this is what I am working on.

Select files, arrange these in the listbox and click "Insert Pictures".

Each picture will automatically be scaled to best full screen fit, considering image proportions, DPI and what have you (this took quite some time to figure out)
Also animation will be added to each of these:

First picture will appear automatically after a short delay
The remaining will fade in with a slightly longer delay.


##Secure Powerpoint comments

This is my very first project in C# - so there might be a lot of things that could be done better... but it gets the jobs done.

I needed a way to clean up comments from powerpoint presentations before sending them to somebody. On the other hand I did not want to lose the comments, which made the default option of stripping comments a no go.

Running this program will create a copy of the selected presentation without comments. It will also backup existing comments to a standard text file for your convenience.