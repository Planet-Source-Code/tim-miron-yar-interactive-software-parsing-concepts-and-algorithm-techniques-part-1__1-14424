<div align="center">

## Parsing Concepts and Algorithm techniques \(PART 1\)


</div>

### Description

Explores and teaches parsing algorithm techniques (PART 1)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\(Tim Miron\) yar\-interactive software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tim-miron-yar-interactive-software.md)
**Level**          |Intermediate
**User Rating**    |4.0 (28 globes from 7 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tim-miron-yar-interactive-software-parsing-concepts-and-algorithm-techniques-part-1__1-14424/archive/master.zip)





### Source Code


<B><FONT SIZE=4><P>Parsing concepts &amp; Parsing Algorithms</P>
</FONT><FONT SIZE=2><P>INTRODUCTION (about this article)</P>
</B><P>&#9;In this article, I will be looking at the concepts of parsing and parser construction. This article is not meant as a tutorial, or a discussion article, but rather a mix of the two. I will explore mostly the most commonly used and simplest (yet efficient) parsing method - Left to Right / Top to bottom parsing. In this example I will include a few examples for use with Visual Basic. In fact, this article is based around parsing using VB, and will review standard string manipulation functions for people who may not be familiar with them. Please note that I am 15 years old, but I&#8217;ve been programming in VB for 3 years and now use C and C++ as well. (I&#8217;d recommend building any complex parsers in C or C++, due to faster processing speeds, although VB.NET has sufficient processing speed for a large array of parsing tasks.)</P>
<P>PARSING IS looking through a string (a &quot;sentence&quot; or set of characters) and interpreting it as commands, or translating it, or basically setting up reactions when certain &quot;sets&quot; of characters are encountered or found.</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>SECTION 1 - String Manipulation</P>
<P>&#9;Ok, first of all, I will run you through a review of string manipulation functions and string parsing methods (InStr, Mid, Left, Right, Ucase, Lcase etc.) if you&#8217;re already familiar with these functions, you can skip this part and proceed to section 2. As you hopefully already know, a string is a set of characters from the ASCII character set, for example &quot;The Quick Brown Fox Jumped Over The Lazy Dog&quot; is a string.</P>
<P>In parsing, you obviously work with these strings a lot. I will discuss the elements of parsing in SECTION2, but this is just to look over how we &quot;look&quot; ata string. First of all, these are the string manipulation functions available in VB (some may only be available in VB6, I&#8217;ve tried to note the ones that are)</P>
<P>*&#9;Mid&#9;-&#9;For getting a set of characters from the middle of a string.</P>
<P>*&#9;InStr&#9;-&#9;For Searching in one string for an occurrence another string (returns the position of the found string).</P>
<P>*&#9;Left&#9;-&#9;For extracting the leftmost character(s) from a string.</P>
<P>*&#9;Right&#9;-&#9;For extracting the rightmost characters from a string.</P>
<P>*&#9;Replace&#9;-&#9;To replace one set of characters with another (VB6 and up only).</P>
<P>*&#9;Len&#9;-&#9;For getting the length of a specified string.</P>
<P>&nbsp;</P>
<P>I wont dive to deeply into how to use these function&#8217;s syntax, you can refer to the VB help file in the Strings category to find a full set of instructions and examples on how to utilize these. Although we WILL be exploring the use of these functions, and at that time I will explain how to use them properly.</P>
<P>&nbsp;</P>
<P>SECTION 2 - Types of parsers</P>
<P>&#9;There are two different types of parsers, and I&#8217;m not talking about the method in which they parse. I will call them Translator Parsers, and Command Parsers (there may be other types but I&#8217;m sticking to these two) Translators are converters, they take the data they are given and output new data. A compiler would be a perfect example of a translator, it takes the computer syntax and translates it into machine code.</P>
<P> A command parser is a parser that actually does something when it interprets a certain &quot;Command&quot; - for a good example would be a script executor, it finds a command in the script, and it does convert the syntax to a string, instead, it executes a command or subroutine)</P>
<P>&nbsp;</P>
<P>SECTION 3 - The FIVE parts of a parsing algorithm</P>
<P>&#9;I have divided parsing algorithms into 5 basic parts to allow you to better understand the process of parsing a string. They are each defined below&#8230;</P>
<P>INPUT - The data that is given to the parser, usually a string. This is the data that the parser will parse through and will work with in the first place.</P>
<P>OUTPUT - The output is what the end product is (and maybe I should of put this part last in the list) it is what we lend up with after the parsing is complete. The output only exists if the parser is meant to conduct some sort of &quot;translation&quot; of the input (translation parsers take data and output other data accordingly, for more info on types of parsers, see section 2). Like if the parser&#8217;s purpose was to reverse all the letters in a sentence then the output of &quot;Hello World&quot; would be &quot;dlroW olleH&quot; - the input was &quot;Hello World&quot; and the output was &quot;dlroW olleH&quot;.</P>
<P>INTERPRETATION - How the parser interpret the input. Does it see it as a set of commands, or as a language to be translated, what does it look for, what is It trying to find, and what will it do when it finds it. (see section 2 if you haven&#8217;t read it already and don&#8217;t get what I&#8217;m talking about here.)</P>
<P>PROCESSING - What the parser does once it interprets the data, for example, in the VB parser, when it finds the string &quot;MsgBox&quot; it knows that it will be displaying a message box, and the process it takes is looking for the message box properties. Then, after finding the properties (Caption, Buttons, Icon etc.) it displays a message box accordingly. Processing can add to the output depending on what it finds, or it can react, like the message box example, and interpret strings as commands.</P>
<P>&nbsp;</P>
<P>SECTION 4 - Constructing a parser&#8230;</P>
<P>&#9;Yeah! We&#8217;re finally past all that boring $hit about parts of a parser and stuff!!! In this section, we&#8217;re gunna build a parser to execute our own message box script. Go into VB and create a new project and a form and place a textbox on the form and a command button, put he buttons Caption to Execute and make the button&#8217;s name &quot;CmdExe&quot; and keep the name of the textbox as the default "Text1".</P>
<P>Now we&#8217;re gunna construct the parsing algorithm&#8230; When writing an algorithm of any sort, its good to figure out what steps the computer will need to take, or in mathematics, what equation the computer will use. In our case here, we&#8217;ll say that in our new script, the code for a message box is MSB followed by the properties in some angel-brackets then the caption in quotes &#8211; something like this&#8230;</P>
<P>MSB&lt;&quot;Hello World!&quot;&gt;</P>
<P>So the first thing we did was figure out what the script might look like (which is a good idea for any type of script or language designing)</P>
<P> Ok, so what are we gunna do to turn this little script into a message box? &#8211; here&#8217;s what&#8230;</P>
<P>  First of all, we need to find the string that tells us to make a message box &#8211; in this case, we&#8217;re looking for &quot;MSB&quot; so here&#8217;s what we put I our code.</P>
<P>&#8216;---------------------------------------------------------------------</P>
<P>Private Sub CmdExe_Click()</P>
<P> Dim CP 'CP will keep track of the</P>
<P> 'Position of the command</P>
<P> </P>
<P>  CP = InStr(1, UCase(Text1).Text, "MSB")</P>
<P>  'ok, now CP will know the position of the word "MSB"</P>
<P>  'note that we used UCase(Text1.Text) which converts the string</P>
<P> 'in text1 to all uppercase so we don&#8217;t have to worry about</P>
<P> 'case sensitivity</P>
<P>End Sub</P>
<P>&#8216;----------------------------------------------------------------------</P>
<P>Now we&#8217;ve found the command we&#8217;re looking for, this type of parsing isn&#8217;t top to bottom parsing, this type is just finding any possible commands. We should check to make sure the script has a &#8216; &lt;&quot; &#8217; and a &#8216; &quot;&gt; &#8216;. So we&#8217;ll do that and if we know they have put it in, we&#8217;ll need to find the caption of the message box, otherwise give them an error message! We&#8217;ll be storing the caption for further use as a variable. We can call our variable &quot;MBCap&quot; so here&#8217;s what the code will look like now&#8230;</P>
<P>&#8216;----------------------------------------------------------------------</P>
<P>Private Sub CmdExe_Click()</P>
<P> Dim CP, CP2, CP3, CP4 'CP will keep track of the</P>
<P> 'Position of the command</P>
<P> Dim MBCap As String 'Stores the caption of</P>
<P> 'the message box for further use</P>
<P> </P>
<P>  CP = InStr(1, UCase(Text1.Text), "MSB")</P>
<P>   If CP = 0 Then Exit Sub 'if we dont find it, discontinue</P>
<P>  'If we found it it will continue</P>
<P>  'ok, now CP will know the position of the word "MSB"</P>
<P>  'note that we used UCase(Text1.Text) which converts the string</P>
<P> 'in text1 to all uppercase so we dont have to worry about</P>
<P> 'case sensitivity</P>
<P>  'NOW WE CHECK FOR THE &lt;" and "&gt;</P>
<P> CP2 = Mid(Text1.Text, CP + 3, 2) 'this selects the 2 characters directly</P>
<P> 'after the word MSB</P>
<P> 'check for the second</P>
<P>   CP3 = InStr(CP + 5, Text1.Text, Chr(34) &amp; "&gt;")</P>
<P>   CP4 = Mid(Text1.Text, CP3, 2)</P>
<P>   </P>
<P>  If CP2 = "&lt;" &amp; Chr(34) And CP4 = Chr(34) &amp; "&gt;" Then</P>
<P>  'the if says if we found &lt;" and "&gt; the continue</P>
<P>   Else</P>
<P>     Exit Sub 'otherwise discontiue with this sub</P>
<P>  End If</P>
<P>End Sub</P>
<P>&#8216;-----------------------------------------------------------------------------</P>
<P> As you more advanced programmers can see, I haven&#8217;t been the most efficient, but this is just one of thoughs things where simple is better. Now we need to extract the caption of the button, so this is how we do that, we&#8217;re gunna find the length of the caption by subtracting the position of the &quot;&gt; from the position of the &quot;&gt;, then we&#8217;ll select the caption and store it as a string for later use.</P>
<P>Now this is what the code should look like&#8230; (Copy it into your program, be sure to study it though)</P>
<P>&#8216;-------------------------------------------------------------------------------</P>
<P>Private Sub CmdExe_Click()</P>
<P> Dim CP, CP2, CP3, CP4 'CP will keep track of the</P>
<P> 'Position of the command</P>
<P> Dim MBCap As String 'Stores the caption of</P>
<P> 'the message box for further use</P>
<P> Dim CapLen As Integer 'stores the captions length</P>
<P> </P>
<P>  CP = InStr(1, UCase(Text1.Text), "MSB")</P>
<P>   If CP = 0 Then Exit Sub 'if we dont find it, discontinue</P>
<P>   'If we found it it will continue</P>
<P>  'ok, now CP will know the position of the word "MSB"</P>
<P>  'note that we used UCase(Text1.Text) which converts the string</P>
<P> 'in text1 to all uppercase so we dont have to worry about</P>
<P> 'case sensitivity</P>
<P>  'NOW WE CHECK FOR THE &lt;" and "&gt;</P>
<P> CP2 = Mid(Text1.Text, CP + 3, 2) 'this selectd the 2 characters directly</P>
<P> 'after the word MSB</P>
<P> 'check for the second</P>
<P>   CP3 = InStr(CP + 5, Text1.Text, Chr(34) &amp; "&gt;")</P>
<P>   CP4 = Mid(Text1.Text, CP3, 2)</P>
<P>   </P>
<P>  If CP2 = "&lt;" &amp; Chr(34) And CP4 = Chr(34) &amp; "&gt;" Then</P>
<P>  'the if says if we found &lt;" and "&gt; the continue</P>
<P>   Else</P>
<P>     Exit Sub 'otherwise discontinue with this sub</P>
<P>  End If</P>
<P>CapLen = CP3 - (CP + 5)</P>
<P> MBCap = Mid(Text1.Text, CP + 5, CapLen)</P>
<P> MsgBox MBCap</P>
<P>End Sub</P>
<P>&#8216;--------------------------------------------------------------------------------</P>
<P>NOW if you run it and type MSB&lt;&quot;Hello WORLD!&quot;&gt; in the textbox, press execute and you get a message box with &#8216;hello WORLD!&#8217; on it!</P>
<P> I got tired hands, and I&#8217;m only 15, and I have school, so I&#8217;ll continue this tomorrow, happy programming, please vote!</P></FONT>

