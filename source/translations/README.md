CSCANPST looks for text in the SCANPST windows to know what it going on, so the text it is looking for must match *exactly* 
what is displayed in the window (is does not understand the text - it is just a letter for letter match). 

This means that words in other languages will not match. Luckly some kind CSCANPST users have made versions to work with their languages.

If you want to make a translation for your language...

Not hard. Just edit the cscanpst.ahk file with a text editor. 

Anyplace you see` FileAppend`, this is a message to the user. The program just outputs whatever you put here, so you can freely
translate these to be informative in yuor language. 

Anyplace you see `ifWinExist`, the text at the end of the line must exactly match the text in your version of SCANPST. 
You can find this text by manually walking thought the SCANPST.EXE program and noting the text, although it might be hard
to know exactly what some of the errors look like if you have never encountered them. Maybe a google search can find them?

Here is an example of a translation so you can see which parts changed (although he did not update the user alert messages so they are still in English)…

https://github.com/bigjosh/CSCANPST/blob/master/source/translations/cscanpst.pt-br.ahk

When you are done, either send me a pull-rqeuest if you are on Github or email me the file and I'll post the changes for others to enjoy. Thanks!
