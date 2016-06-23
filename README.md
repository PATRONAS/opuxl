# opuxl
Excel Addin to connect your favorite programming language with your Excel sheets

The Opuxl Excel Addin allows you to connect your application with Excel. You define your functions in your favorite programming language (Java, Node, ...) and the Opuxl Addin can call those functions from within your Excel Sheet.

We had trouble to fill in multiple cells via a single function by using [XLLoop](http://xlloop.sourceforge.net/), so we created our own addin which uses [ExcelDNA](https://exceldna.codeplex.com/) to manipulate the Excel sheet.


## How does it work?
When you start your Excel with the plugin activated, it will try to create a socket connection to 127.0.0.1:61379 and send a "Initialize" request to fetch all available methods. The response will contain the methods and they will be registered in the Excel Sheet. A method invocation will trigger a stateless request/response connection between the Addin and the "Socket Server".

Check out the Java Demo which creates a Server Socket and published methods which have a specific annotation.

(Watch out: Currently all responses from the server socket should be a matrix, eg: Object[][], which is directly printed into the Excel sheet) 

## How to connect the demo Java server with your Excel sheet

**Java Part with Eclipse**

1. Import "demos/java" as a Maven Project
2. Project -> Maven -> Update Project
3. Run the DemoServer
(this will publish all methods of the DemoMethods Class which are annotated with @OpuxlMethod)

Now the Socket Server is running and we have to start an Excel with an installed addin.

**Debug mode with Visual Studio Community 2015**

1. Import "opuxl_addin/Opuxl" as a Solution
2. Install NuGet Packages on both Projects
3. OpuxlClassLibrary Project -> Properties -> Debug
...Set 'Start External Program' to your EXCEL.exe file
...Add a 'Commandline Argument' to your xll file (eg "C:\Users\foo\Documents\Visual Studio 2015\Projects\Opuxl\OpuxlClassLibrary\bin\Debug\Opuxl.xll")
4. Start the Project in 'debug' mode
5. Execute your functions defined in your java demo server
...they are prefixed with "Opus." in this demo build, eg: =Opus.GetNames()

You can find further information inside the OpuxlClassLibrary Project folder: [opuxl_addin/Opuxl/OpuxlClassLibrary](https://github.com/PATRONAS/opuxl/tree/master/opuxl_addin/Opuxl/OpuxlClassLibrary)

### Want to help? Or found a Bug? Or a specific feature is missing? Or just a question?
Excellent! You can help us by just creating an Issue or submiting a Pull Request. Your help would be appreciated.

## Example
![Example](/opuxl-demo.gif?raw=true "")
