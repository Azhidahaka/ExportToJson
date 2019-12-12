----------------------------------------------
            Version 1.0.1
----------------------------------------------
v.1.0.1
Folder Structure Change
Key Image Change

v.1.0.0
Release ExcelToJson Package

———————————————————
 Supprot
———————————————————
c#    : .net 2.0
unity : 5.0 over And PC, Mac & Linux Standalone 
excel : xls, xlsx

———————————————————
 Menu
———————————————————
1. Window/Open ExcelToJsonEditor
2. Assets/Convert ExcelToJson

———————————————————
 tutorials  
———————————————————
Example Description

Asset folder
Sample/ExcelData : Excel file storage,  Excels2JsonEditor ->Load Excel Folder Change 
Sample/Script	 : Sample Excel Read Source
Json	         : Output Json files 
ExcelToJson      : Source Code  (Do not delete) 

0. Common Tutorials
-> Open ExcelToJsonSample.Scene
-> Play 
-> View Debug Log

1. Simple Convert 
-> Select CustomData.xls, SimpleData In Excel folder, 
-> right-click and Convert ExcelToJson 
-> Editor All Convert 

2. View SimpleData.xls

3. View CustomData.xls 
-> Following Grammer

4. Window/Open Excel2JsonEditor
-> You are Excel Folder Select
-> All Conver Or Sheet Convert 

———————————————————
 CusomJson Excel Grammar
———————————————————
JSON_OUTPUT		 : excel file -> json output name
JSON_VERSION	 : json convert version (Ex 1.0.0)
JSON_START  	 : json object start
JSON_END   		 : json object end
JSON_ARRAY_START : json array start (if you use this, It should put the name of the array
JSON_ARRAY_END   : json array end
JSON_TABLE_START : json table start (if you use this, It should put the name of the table)
JSON_TABLE_END	 : json table end	
JSON_VALUE       : json value (if you use this, you can string, int, float, bool, null)

ex) Excel/CustomData.xls