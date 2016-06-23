# opuxl
Excel Addin to connect your favorite programming language with your Excel sheets

The Opuxl Excel Addin allows you to connect your application with Excel. You define your functions in your favorite programming language (Java, Node, ...) and the Opuxl Addin can call those functions from within your Excel Sheet.

We had trouble to fill in multiple cells via a single function by using [XLLoop](http://xlloop.sourceforge.net/), so we created our own addin which uses [ExcelDNA](https://exceldna.codeplex.com/) to manipulate the Excel sheet.


## How does it work?
When you start your Excel with the plugin activated, it will try to create a socket connection to 127.0.0.1:61379 and send a "Initialize" request to fetch all available methods. The response will contain the methods and they will be registered in the Excel Sheet. A method invocation will trigger a stateless request/response connection between the Addin and the "Socket Server".

Check out the Java Demo which creates a Server Socket and published methods which have a specific annotation.

(Watch out: Currently all responses from the server socket should be a matrix, eg: Object[][], which is directly printed into the Excel sheet) 


## Want to help? Or found a Bug? Or a specific feature is missing?
Excellent! You can help us by just creating an Issue or submiting a Pull Request. Your help would be appreciated.

## Example
![Example](/opuxl-demo.gif?raw=true "")
