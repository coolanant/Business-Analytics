# Excel-Crash-Course

<details close>
  <summary><b><i>Basics </i></b></summary>

1. Quick Access Toolbar
2. Ribbon Menu
3. Shortcut Menu/ Mini Toolbar

Every **workbook**(file) has atleast one **worksheet**<br/>
Every worksheet has same no. of row and col<br/>
Numbers/formula/date are right aligned<br/>

* Fill Pointer
* Merge and center
* Insert/delete/hide/unhide
* Move/copy/cut
* pivot table
* Custom Data Types/ power BI
* Security
* Sharing/Tracking

### Date and Time

### Shortcuts
* Auto Sum - **alt =**
* Format Cells - **ctrl 1**
* Chart - **alt f1** / **alt f11**
* Edges - **Cntrl . **

</details>

<hr/>

# Advanced

* Show Formulas - **Ctrl** ~
* Solve part of formula - **f9**
<hr/>

* Auditing Tools
* Trace Dependendents
* Trace Precedents
* add/multiply values to all
* define name/var (f3)



<hr/>

### Formulas
* =sum(b2:G2)
* =averge()
* =max()
* =min()

* =subtotal() -9 sum... (ignores other subtotals)

<br/>

* =xlookup()
* =vlookup(x, table,tablerowno.) - ascending order
* =hlookup()

<br/>

* =trim() - remove leading and trailing spaces
* =proper() - capitalize 1st letter

<br/>

* =countif(F, "Full Time")
* =averageif(F, "Full Time")
* =sumif(F, "* Time")

<br/>

* = countifs(F, "Full time",D,"part time", E,">9")
* maxifs()
* minifs()

<br/>



<hr/>

### IF
* if(h2>3, 3000,0)
* if(h2>3, 3000, if())
* if(and(x,y), c, d)
* if(or(x,y), c, d)
* and, or, not

<br/>

*  ifs(h1=20,0, h1=30,1, h1=40,2, true,3)
