# RecursiveFunctionVB6
Recursive Function - Building XML Categories Structure

ADO is used to to loop through tables in database dbRecursiveCategory. Find all categories (main and child) and create a XML file.
Featured Highlights:

The recursive function will find all children categories and will transform them into XML file. There is no limitation in category depth.

Requirements:
Visual Basic 6.00 with SP6
SQL Server 2000 or Microsoft Data Engine (MSDE)
Microsoft ActiveX Data Objects 2.8 Library
Microsoft XML, v4.0

Running the Sample:
Restore database located into folder Db_Sql2000
Execute function CreateChildrenFunction.sql located into folder SqlDataScripts
or generate tables and import data from folder SqlDataScripts Execute CreateTablesQuery.sql and CreateChildrenFunction.sql Import tables from files: tbCategoryMain.tbl, tbCategoryRelation.tbl, tbCategoryText.tbl
Modify the Connection string: variable STR_CONNECT and recompile application
The XML will be located in the executable folder - RecursiveFunc.xml


See Also
=========================================================================================
Author: Kriss Nickov
Project web site: 	http://www.kncode.us/projects
Author Resume: 		http://www.kncode.us/profile
Published:			February, 25 2917
=========================================================================================