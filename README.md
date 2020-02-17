# Excel-Hyperlinks
The main purpose of this script was to save time.
A workbook was created and it was later decided to start linking emails within the excel workbook.
There were hundreds of excel entries that would have to have been manually hyperlinked in excel by many people.

So instead I created a script that would search within your file directory and link the file based on how it's been named.  

Currently it works by having the excel file and all the files you would like to link in the same directory.  But if these are in different locations, all you would need to do is change the file directory in the script.

A quick view of what our data initally looks like.

![rawdata](https://user-images.githubusercontent.com/23482152/74617869-5d456b80-50fd-11ea-849d-7aa83d545e84.png)

And a quick look at what our file folder looks like.  More importantly, we want to take a look at how these files are named.  How they are named will play a big part in how effective our script is.

![filenames](https://user-images.githubusercontent.com/23482152/74617895-7bab6700-50fd-11ea-894b-ab7e993bfed9.png)

At the time, there is no control for how these files are named, so inevitably there will be cases where the files are not linked.  But as of now, with our tests, we have linked all that we have in our file folder.  Any unlinked cells are due to the file not existing in the folder.

So in the end, this is what we want to end up with.  By clicking on the highlighted cell, we will be taken directly to the file and it will open up for us to reference.

![finalview](https://user-images.githubusercontent.com/23482152/74617949-b6ad9a80-50fd-11ea-9f7a-fa39c0182c9b.png)


