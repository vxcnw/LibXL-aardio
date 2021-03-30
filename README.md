# libxl-bindings-for-aardio
#### [libxl](https://www.libxl.com/) LibXL is a library that can read and write Excel files. It doesn't require Microsoft Excel and .NET framework, combines an easy to use and powerful features. Library can be used to
* Generate a new spreadsheet from scratch
* Extract data from an existing spreadsheet
* Edit an existing spreadsheet
#### [aardio](http://www.aardio.com/)  is an extremely easy-to-use dynamic language, but it is also a hybrid language that allows for rare and very convenient manipulation of static types, so it can directly call API interface functions of static languages such as C, C++, etc.

# different frome LibXL
* Index starts at 1

## example
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

````javascript
io.open()

import aaz.libxl;

var book = aaz.libxl.createBook();
book.setKey();
book.load("\example.xls");

var sheet = book.getSheet(1);   

for(i=sheet.firstRow; sheet.lastRow; 1){
	for(j=sheet.firstCol; sheet.lastCol; 1){
		var ret;
		select(sheet.cellType(i, j) ) {
			case 0/*CELLTYPE_EMPTY*/ {
				io.print(i, j , "空")
			}
			case 1/*CELLTYPE_NUMBER*/ {
				ret = sheet.readNum(i, j)
				io.print(i, j ,"数字", ret)
			}
			case 2/*CELLTYPE_STRING*/ {
				ret = sheet.readStr( i, j )
				io.print(i, j ,"字符串", ret)
			}
		}
	}
}

book.release()
execute("pause")
````

````javascript
import aaz.libxl;

var book = aaz.libxl.createBook();
book.setKey();

var font = book.addFont();
font.
	setName("Impact").
	setSize(36);

var format = book.addFormat()
format.config = {
    alignH = 2;
    border = 12;
    borderColor = 2;
    font = font;
}

var sheet = book.addSheet("Custom");
sheet.writeStr(2, 1, "Format", format);
sheet.setCol(1, 1, 50);

book.save("\format.xls");
book.release();
````

````javascript
import aaz.libxl;

var book = aaz.libxl.createXmlBook();
book.setKey()

var boldFont = book.addFont(0);
boldFont.bold = 1;

var titleFont = book.addFont(0);
titleFont.name = "Arial Black"
titleFont.size = 16;

var titleFormat = book.addFormat();
titleFormat.font = titleFont;

var headerFormat = book.addFormat();
headerFormat.alignH = 2/*_ALIGNH_CENTER*/
headerFormat.border = 1/*_BORDERSTYLE_THIN*/
headerFormat.font = boldFont;
headerFormat.fillPattern = 1 /*_FILLPATTERN_SOLID*/
headerFormat.patternForegroundColor = 47 /*COLOR_TAN*/

var descriptionFormat = book.addFormat();
descriptionFormat.borderLeft = 1 /*BORDERSTYLE_THIN*/

var amountFormat = book.addFormat();
amountFormat.numFormat = 5
amountFormat.borderLeft = 1
amountFormat.borderRight = 1

var totalLabelFormat = book.addFormat();
totalLabelFormat.borderTop = 1
totalLabelFormat.alignH = 3
totalLabelFormat.font = boldFont

var totalFormat = book.addFormat();
totalFormat.numFormat = 5
totalFormat.border = 1
totalFormat.font = boldFont
totalFormat.fillPattern = 1
totalFormat.patternForegroundColor = 13

var signatureFormat = book.addFormat();
signatureFormat.alignH = 2
signatureFormat.borderTop = 1

var sheet = book.addSheet( "Invoice" )
sheet.writeStr(2, 1, "Invoice No. 3568", titleFormat)

sheet.writeStr(4, 1, "Name: John Smith")
sheet.writeStr(5, 1, "Address: San Ramon, CA 94583 USA")

sheet.writeStr(7, 1, "Description", headerFormat)
sheet.writeStr(7, 2, "Amount", headerFormat)


sheet.writeStr( 8, 1, "Ball-Point Pens", descriptionFormat);
sheet.writeNum(8, 2, 85, amountFormat);
sheet.writeStr( 9, 1, "T-Shirts", descriptionFormat);
sheet.writeNum(9, 2, 150, amountFormat);
sheet.writeStr( 10, 1, "Tea cups", descriptionFormat);
sheet.writeNum(10, 2, 45, amountFormat);

sheet.writeStr( 11, 1, "Total:", totalLabelFormat);
sheet.writeFormula(11, 2, "=SUM(C9:C11)", totalFormat);

sheet.writeStr(14, 2, "Signature", signatureFormat);

sheet.setCol( 1, 1, 40, null, 0);
sheet.setCol(2, 2, 15, , 0);

book.save("\invoice.xlsx")
book.release()
````
