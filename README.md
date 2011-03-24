CP-1252 Fixer
=============

About
-----
You work with the internet.  
Your job involves copying and pasting content from Microsoft Word.  
Funny characters like **&acirc;&euro;&trade;**, **&acirc;&euro;&tilde;** and **&acirc;&circ;&Agrave;&circ;&Ugrave;** show up on your webpages.

This is a small program which runs in your system tray and monitors your clipboard for the presence of "smart quotes" and other characters normal HTML documents may have problems with. These are automatically fixed without any action on the user's part, leaving you able to go about your work error-free.

Features
--------
* Keep your webpages safe from potentially damaging markup
* Various levels of escaping, all independently toggleable:
	* Replace smart quotes with ASCII equivalent characters (**&lsquo; -> '** etc)
	* Replace smart quotes with HTML entities, to retain their presentation safely (**&lsquo; -> `&lsquo;`** etc)
	* Replace any characters with named entity equivalents. This keeps most potentially unsafe characters from *any* source appearing incorrectly in ASCII output. (replaces any characters listed [here](http://www.w3.org/TR/html40/sgml/entities.html))
	* Replace all high ASCII characters with entity number references. This will ensure maximum compatibility by replacing any nonstandard characters at all. This is useful for converting Asian texts to HTML content, but would obviously not be desirable when working with internationalised character sets.
	* Optionally convert rich text data to plain text. This is useful if you wish to remove formatting from copied text. By default, rich text data copied is retained along with its original formatting.
* Persistent settings

Explanation
-----------
Microsoft Word, Outlook, Powerpoint and some other Windows programs use a feature called "smart quotes", which is designed to automatically substitute different characters for keys on your keyboard as you type. This is in an attempt to promote readability via the use of typographer's quotes and punctuation marks. As a result:

<table style="font-family: monospace">
    <tr>
        <th>When you type...</th>
        <th>this appears:</th>
    </tr>
    <tr><td> 'text </td><td> &lsquo;text </td></tr>
    <tr><td> text' </td><td> text&rsquo; </td></tr>
    <tr><td> "text </td><td> &ldquo;text </td></tr>
    <tr><td> text" </td><td> text&rdquo; </td></tr>
    <tr><td> - text </td><td> &bull; text </td></tr>
    <tr><td> -- </td><td> &mdash; </td></tr>
    <tr><td> - </td><td> &ndash; </td></tr>
    <tr><td> ... </td><td> &hellip; </td></tr>
</table>

When pasted into a *normal* HTML document [encoded as ASCII text](http://en.wikipedia.org/wiki/ISO-8859-1#ISO-8859-1), the character set of the document doesn't know how to display these characters. Instead, they are spit out as the combination of ASCII letters which make up the longer number of the extended character, resulting in broken text. For example, &lsquo; is `0xE28098`:

<pre>
           <strong>0xE28098</strong>
         = 0xE2 0x80 0x98
in ASCII:  &acirc;    &euro;    &tilde;
</pre>

These problems can be avoided by:

* Disabling "smart quotes" in Word (which doesn't help unless you are the original author of the document)
* Using a UTF-8 encoded webpage (which doesn't help unless you have control over the website's design)
* Running search/replace algorithms in your text editor before uploading (which is damn time consuming!)

However even then, you should keep your text portable! Anyone copying these characters off your webpage onto a Linux or Mac computer will not be able to read them reliably, as they are Windows - only characters and have no equivalents or *goddamn place* elsewhere in the digital world. Seriously. Even Microsoft's own file formats encode them as escape sequences. Guys.

License
-------
This software is provided under an MIT open source license, read the 'LICENSE.txt' file for details.

Known Issues
------------
* Seems to lose its handle on the clipboard after a computer hibernation, probably needs to intercept an ACPI event and re-register?
