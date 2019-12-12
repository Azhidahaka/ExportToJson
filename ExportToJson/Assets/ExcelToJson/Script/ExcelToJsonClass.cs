/**
 * @brief Having information and the ability to convert Excel data into json
 */

namespace ExcelToJson
{
	#region namespace
	using UnityEngine;
	using UnityEditor;
	using System.Collections;
	using System.Collections.Generic;
	using NPOI.SS.UserModel;
	using NPOI.HSSF.UserModel;
	using NPOI.XSSF.UserModel;
	using System.IO;
	using JsonFx;
	using SimpleJSON;
	using System;
	using System.Text;
	#endregion

	/**
	 * @brief Excel To Json convert function
	 */
	#region excelToJson Convert Function
	public class ExcelToJsonConvert
	{
		public static void ConvertToJson(Sheet sheet, string outputPath)
		{
#if !UNITY_WEBPLAYER 
			JSONClass json = GetJson(sheet.jsonFieldList);

			if(json == null) return;
		
			string filePath = outputPath + "/" + sheet.outputName + ".json";

			FileInfo fileInfo = new FileInfo(filePath);
			if(fileInfo.Exists)
				fileInfo.Delete();

			FileStream fs = fileInfo.OpenWrite();

			if(fs != null)
			{
				StreamWriter sw = new StreamWriter(fs);
				
				if(sw != null)
				{
					string data = json.ToString();
					sw.Write(data);
					
					sw.Close();
				}
				fs.Close();
			}

			AssetDatabase.Refresh();
#endif
		}

		public static JSONClass GetJson(List<JsonField> field)
		{
			int index = 0;
			JSONClass json = (JSONClass)ToJson(field, ref index, null);
			return json;
		}

		static JSONNode ToJson(List<JsonField> field, ref int index, JSONNode node)
		{
			if(index >= field.Count) return null;

			JSONNode result = null;

			if(field[index].jsonType == JSON_TYPE.JSON_START)
			{
				JSONClass json  = new JSONClass();

				if(node != null)
				{
					if(!string.IsNullOrEmpty(field[index].value))
						node.Add(field[index].value, json);
					else
						node.Add(json);
				}
				do
				{
					index++;
					result =  ToJson(field, ref index, json);

				}while(result != null);

				return json;
			}
			else if(field[index].jsonType == JSON_TYPE.JSON_ARRAY_START || field[index].jsonType == JSON_TYPE.JSON_TABLE_START)
			{
				JSONArray array = new JSONArray();

				if(!string.IsNullOrEmpty(field[index].value))
					node.Add(field[index].value, array);
				else
					node.Add(array);
				
				do
				{
					index++;
					result = ToJson(field, ref index, array);
					
				}while(result != null);
				
				return array;
			}
			else if(field[index].jsonType == JSON_TYPE.JSON_VALUE)
			{
				node.Add(field[index].fieldName, field[index].value);
				return node;
			}
			else if(field[index].jsonType == JSON_TYPE.JSON_END || field[index].jsonType == JSON_TYPE.JSON_ARRAY_END ||
			        field[index].jsonType == JSON_TYPE.JSON_TABLE_END )
			{
				return null;
			}

			return null;
		}
	}
	#endregion

	/**
	 * @brief Class that manages the data of the file as a whole excel
	 */
	#region ExcelInfo
	public class Book
	{
		public ExcelToJson.Sheet[] sheet;				/**<< The information in Excel sheet 						   								*/
		public string 	   filePath;					/**<< Excel file path, that path will vary if the Data Deleted 								*/
		public string 	   fileName;					/**<< The name of the Excel file 							   								*/
		public bool 	   foldOut = true;				/**<< As seen on the editor determines whether the exposure of the entire sheet information */ 
		public bool 	   isData;						/**<< If you load an Excel file information													*/
		
		static string[] _extensionInfo 	  = {".xls", ".xlxs"};

		public void SetBook()
		{
			if(string.IsNullOrEmpty(filePath))
			{
				isData = false;
				return;
			}
			
			isData = true;
			IWorkbook book = GetWorkBook();
			if (book == null) {
				isData = false;
				return;
			}
			
			sheet = new Sheet[book.NumberOfSheets];
			for(int i=0; i<book.NumberOfSheets; i++)
			{
				Sheet sheetInfo = new Sheet();
				sheetInfo.SetSheet(fileName, book.GetSheetAt(i));
				sheet[i] = sheetInfo;
			}
		}
		
		IWorkbook GetWorkBook()
		{
#if !UNITY_WEBPLAYER 
			FileInfo fileInfo = new FileInfo(filePath);
			if(!fileInfo.Exists) return null;

			try
			{
				FileStream stream = fileInfo.OpenRead();
				if(stream == null)
				{
					stream.Close();
					return null;
				}

				fileName = GetFileName(filePath);

				IWorkbook book;
				if(fileName.Contains("xlsx"))
					book = new XSSFWorkbook(stream);
				else 
					book = new HSSFWorkbook(stream);

				return book;
			}
			catch(IOException e) {
				CustomEditorTools.DisplayNoticeDialog ("Error", "Close Excel File And Try Again\n" + e.Message);
			}
#endif
			return null;
		}
		
		public static string GetFileName(string filePath)
		{
			string[] fileSplit = filePath.Split('/');
			
			for(int i=0; i<fileSplit.Length; i++)
			{
				for(int j=0; j<_extensionInfo.Length; j++)
				{
					if(fileSplit[i].Contains(_extensionInfo[j]))
					{
						string[] finalFileName = fileSplit[i].Split('.');
						return finalFileName[0];
					}
				}
			}
			
			return "";
		}
	}
	#endregion

	/**
	 * @brief Information for converting the field and value information in Excel as Json
	 */
	#region jsonfield
	public class JsonField
	{
		public JSON_TYPE   jsonType;
		public string      value;
		public string 	   fieldName;

		public JsonField()
		{
			jsonType  = JSON_TYPE.JSON_START;
			value	  = "";
			fieldName = "";
		}
	}
	#endregion

	/**
	 * @brief Class to manage the information in Excel sheet
	 */
	#region SheeInfo
	public class Sheet
	{
		public string 	sheetName;						 	 /**< Name of the Excel sheet 											 */
		public string 	outputName;							 /**< Name information when converts json								 */
		public string 	version;
		ISheet		    _sheet;

		public List<ExcelToJson.JsonField> jsonFieldList	 = new List<JsonField>();

		public bool IsSheet()
		{
			if(_sheet == null) return false;

			return true;
		}

		/**< When you first read the data from the sheet sets the default value */
		public void SetSheet(string bookName, ISheet sheet)
		{
			_sheet 		 = sheet;
			sheetName 	 = _sheet.SheetName;
			InitFieldInfo();
			if (string.IsNullOrEmpty (outputName))
				outputName = sheetName;
		}

		/*< When you import the information in the sheets must be reloaded again the value of a field */
		public void CopySheet(ISheet sheet)
		{
			_sheet = sheet;
			InitFieldInfo();
		}

		public void AddFieldInfo(JSON_TYPE type, string fieldName, string value)
		{
			JsonField field = new JsonField();
			field.jsonType  = type;
			field.value     = value;
			field.fieldName = fieldName;
			jsonFieldList.Add (field);
		}

		void InitFieldInfo()
		{
			jsonFieldList.Clear();
			AddFieldInfo(JSON_TYPE.JSON_START, "", "");
			LoadFieldInfo();
			AddFieldInfo(JSON_TYPE.JSON_END, "", "");
		}

		/**< how to read json as table , How to read from Excel into jsontype */ 
		void LoadFieldInfo()
		{
			JSON_TYPE type = GetJsonType(0, JSON_TYPE.ERROR);
			int index 	   = 0;
			if (type == JSON_TYPE.JSON_OUTPUT)
			{
				IRow rowInfo = _sheet.GetRow(0);
				ICell cellField = rowInfo.GetCell(1);
				if(cellField != null) 
					outputName =  cellField.StringCellValue;
				
				index 	   = 1;
				type = GetJsonType (1, JSON_TYPE.ERROR);
			}
			if (type == JSON_TYPE.JSON_VERSION)
			{
				IRow rowInfo = _sheet.GetRow(index);
				ICell cellField = rowInfo.GetCell(1);
				if(cellField != null) 
					version =  cellField.StringCellValue;

				index 	   = index+1;
				type = GetJsonType (index, JSON_TYPE.ERROR);
				AddFieldInfo (JSON_TYPE.JSON_VALUE, "Version", version);
			}
			if(type == JSON_TYPE.ERROR)
			{
				AddFieldInfo(JSON_TYPE.JSON_TABLE_START, "", sheetName);
				LoadFieldInfoByExcel(JSON_TYPE.JSON_TABLE_START, index);
				AddFieldInfo(JSON_TYPE.JSON_TABLE_END, "", "");
			}
			else
			{
				AddFieldInfo(JSON_TYPE.JSON_START, "RES_NAME", sheetName);
				LoadFieldInfoByExcel(JSON_TYPE.ERROR, index);
				AddFieldInfo(JSON_TYPE.JSON_END, "", "");
			}
		}

		/*<< Having read the sheet by one line, convert jsontype */ 
		JSON_TYPE GetJsonType(int rowIndex, JSON_TYPE lastType)
		{
			if(_sheet == null) return JSON_TYPE.ERROR;

			IRow rowInfo   = _sheet.GetRow(rowIndex);
			if(rowInfo == null) return JSON_TYPE.ERROR;
			
			ICell CellInfo = rowInfo.GetCell(0);
			
			if(CellInfo == null) return JSON_TYPE.ERROR;
			
			string jsonType ="";
			
			if(CellInfo.CellType == CellType.Numeric)
				jsonType = CellInfo.NumericCellValue.ToString();
			else
				jsonType = CellInfo.StringCellValue;
			
			if(string.IsNullOrEmpty(jsonType)) return JSON_TYPE.ERROR;
			
			jsonType	= jsonType.ToUpper();
			
			foreach (JSON_TYPE type in Enum.GetValues(typeof(JSON_TYPE)))
			{
				if(jsonType.Equals(type.ToString())) return type;
			}
			
			if(lastType == JSON_TYPE.JSON_TABLE_START)
			{
				return JSON_TYPE.JSON_TABLE_FIELD;
			}
			if(lastType == JSON_TYPE.JSON_TABLE_FIELD || lastType == JSON_TYPE.JSON_TABLE_VALUE)
			{
				return JSON_TYPE.JSON_TABLE_VALUE;
			}
			
			return JSON_TYPE.ERROR;
		}

		bool LoadFieldInfoByExcel(JSON_TYPE lastType, int index = 0)
		{
			if(_sheet == null) return false;

			List<string> tableField = new List<string>();

			for(int row=index; row<=_sheet.LastRowNum; row++)
			{
				JSON_TYPE type = GetJsonType(row, lastType);

				if(type != JSON_TYPE.ERROR) 
					lastType = type;

				IRow rowInfo = _sheet.GetRow(row);

				if(rowInfo == null) return false;

				if(type == JSON_TYPE.JSON_ARRAY_START 		|| type == JSON_TYPE.JSON_ARRAY_END || type == JSON_TYPE.JSON_TABLE_END ||
				   type == JSON_TYPE.JSON_START		  	    || type == JSON_TYPE.JSON_END)
				{
					string value = "";

					ICell cellField = rowInfo.GetCell(1);
					if(cellField != null) 
						value =  cellField.StringCellValue;
					else
					{
						if(type == JSON_TYPE.JSON_ARRAY_START || type == JSON_TYPE.JSON_TABLE_START)
						{
							CustomEditorTools.DisplayNoticeDialog("Warning Sheet : " + sheetName, 
							                                      "table and field names array is required. If the field name is not difficult to use and accessible to JSON.");
						}
					}

					AddFieldInfo(type, "", value);
				}

				if(type == JSON_TYPE.JSON_VALUE)
				{
					ICell cellField = rowInfo.GetCell(1);
					if(cellField == null) return false;

					string field = cellField.StringCellValue;

					ICell cellValue = rowInfo.GetCell(2);

					if(cellValue == null) return false;

					if(cellValue.CellType == CellType.Numeric)
						AddFieldInfo(JSON_TYPE.JSON_VALUE, field, cellValue.NumericCellValue.ToString());
					else
						AddFieldInfo(JSON_TYPE.JSON_VALUE, field, cellValue.StringCellValue);
				}

				if(type == JSON_TYPE.JSON_TABLE_START)
				{
					ICell cell = rowInfo.GetCell(1);
					if(cell == null) return false;
					AddFieldInfo(type, "", cell.StringCellValue);
				}

				if(type == JSON_TYPE.JSON_TABLE_FIELD)
				{
					for(int col=0; col<rowInfo.LastCellNum; col++)
					{
						ICell cell = rowInfo.GetCell(col);
						if(cell == null) return false;
						tableField.Add(cell.StringCellValue);
					}
				}

				if(type == JSON_TYPE.JSON_TABLE_VALUE)
				{
					AddFieldInfo(JSON_TYPE.JSON_START, "", "");
					for(int col=0; col<rowInfo.LastCellNum; col++)
					{
						ICell cell = rowInfo.GetCell(col);
						if(cell == null) return false;

						if(cell.CellType == CellType.Numeric)
							AddFieldInfo(JSON_TYPE.JSON_VALUE, tableField[col], cell.NumericCellValue.ToString());
						else
							AddFieldInfo(JSON_TYPE.JSON_VALUE, tableField[col], cell.StringCellValue);
					}

					AddFieldInfo(JSON_TYPE.JSON_END, "", "");
				}
			}

			return true;
		}
	}
	#endregion
}
