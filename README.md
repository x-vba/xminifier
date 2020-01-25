# XMinifier

## Description

[XMinifier](http://x-vba.com/xminifier) is a utility tool used to minify your VBA source code by removing
excess comments, whitespace, and syntax features that don't have an effect on
the actual functioning of your code. This can be useful when you want to ship
your code or include it into a Document, Spreadsheet, or Presentation, as the
minified source code reduce the total size of the code by 50% or more. In one
case, I was able to minify my code by almost 70%, making it much smaller when
sending the Spreadsheet to others and when shipping my code.

## Usage

XMinifier is written in pure ES6 JavaScript, which means it can be run in the
browser and does not require any external downloads to run. To use it, simply
go to the XMinifier web page and follow the prompts.

## Where does my code go?

Since XMinifier is written in pure ES6 JavaScript and has no external
dependencies, your code is never shipped to a server to be minified. XMinifier
is run purely locally and can be run offline by saving the web page. Additionally,
the source code for XMinifier can be found in the web page in unminified form, 
so you can be sure that your VBA code remains with you.

## XMinifier isn't running?

XMinifier is written in ES6 JavaScript. Some older browsers don't support ES6
JavaScript (notably older versions of Internet Explorer). If XMinifier does not
run in your browser, try using a different browser, such as Chrome, Firefox, or
Safari.

## License

The MIT License (MIT)

Copyright © 2020 Anthony Mancini

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 
