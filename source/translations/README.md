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


These words show up in the windows when you run SCANPST by itself. For example, when I run scan PST I see…

![](images/repair.gif)

The CSCNAPST program is looking for the words “Microsoft Inbox Repair Tool“  in the header of the windows.  Yours might say something different – in French presumably. 

When I run a scan on a slightly corrupted file, I see…

![](images/minor.gif)

Here, CSCANPST is looking for the words “Only minor inconsistencies” inside the window. See?

You can pick any words inside the target window as long as they are not in any other window. It is nice, however, to pick words the help explain which window you are looking for, so “in this file. Repairing the” would also work on my version, but not as clear which window that refers to in the code. 




Here is an example of a translation so you can see which parts changed (although he did not update the user alert messages so they are still in English)…

https://github.com/bigjosh/CSCANPST/blob/master/source/translations/cscanpst.pt-br.ahk

When you are done, either send me a pull-rqeuest if you are on Github or email me the file and I'll post the changes for others to enjoy. Thanks!
