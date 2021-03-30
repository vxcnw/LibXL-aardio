# libxl-bindings-for-aardio
#### [libxl](https://www.libxl.com/) LibXL is a library that can read and write Excel files. It doesn't require Microsoft Excel and .NET framework, combines an easy to use and powerful features. Library can be used to
* Generate a new spreadsheet from scratch
* Extract data from an existing spreadsheet
* Edit an existing spreadsheet
#### [aardio](http://www.aardio.com/)  is an extremely easy-to-use dynamic language, but it is also a hybrid language that allows for rare and very convenient manipulation of static types, so it can directly call API interface functions of static languages such as C, C++, etc.

### example
````javascript
import aaz.libxl;

var book = aaz.libxl.createBook()
book.setKey();

var f = {};
var format = {};
var customNumFormats = {
    "0.0";
    "0.00";
    "0.000";
    "0.0000";
    "#,###.00 $";
    "#,###.00 $[Black][<1000];#,###.00 $[Red][>=1000]"
}

for(i=1;#customNumFormats;1){
    f[i] = book.addCustomNumFormat(customNumFormats[i]);
}

for(i=1;#customNumFormats;1){
	format[i] = book.addFormat();
	format[i].numFormat = f[i];
}

var sheet = book.addSheet( "Custom formats" );
sheet.setCol(1, 1, 20, null, 0);

sheet.writeNum(3, 1, 25.718, format[1]);
sheet.writeNum(4, 1, 25.718, format[2]);
sheet.writeNum(5, 1, 25.718, format[3]);  
sheet.writeNum(6, 1, 25.718, format[4]);

sheet.writeNum(8, 1, 1800.5, format[5]);

sheet.writeNum(10, 1, 500, format[6]);
sheet.writeNum(11, 1, 1600, format[6]);

book.save("\custom.xls");
book.release();
````
