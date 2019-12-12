/**
 * @brief Inspector for the Excel file and select Convert Json
 * 		  All data is managed as Json, are stored in the excelToJson folder.
 */

namespace ExcelToJson
{
	using UnityEngine;
	using System.Collections;
	using System.Collections.Generic;
	using UnityEditor;
	using JsonFx;
	using System.IO;
	using System.Text;

	public class ExcelToJsonEditor : EditorWindow
	{
		static readonly string 	 _dirConvert			  = "JsonData";
		static readonly string[] _readExtention		  	  = {"xls", "xlsx"};				/**< The extension of the Excel file to be read from a folder									*/
		static readonly string 	 _excelFolder		 	  = "ExcelToJson/Sample/ExcelData";		/**< The first path is stored in an Excel file													*/
		static readonly string 	 _jsonFolder			  = "ExcelToJson/Json";				/**< The first path is stored in an Json file													*/
		Vector2 				 _scrollPos 	 	  	  = Vector2.zero;

		[SerializeField]
		static List<Book> 		 _currentBookInfo	  = new List<Book>();		/**< Variable that contains information on the current screen shown in Excel 					*/

		[MenuItem("Window/Open Excels2JsonEditor")]
		static public void OpenXls2Json()
		{
#if !UNITY_WEBPLAYER
			EditorWindow.GetWindow<ExcelToJsonEditor>("Excels2JsonEditor");
#else
			CustomEditorTools.DisplayNoticeDialog ("Error", "The editor does not work on the Web Player environment. We recommend a PC environment.");
			return;
#endif
		}

		[MenuItem("Assets/Convert ExcelToJson")]
		static public void ConvertXls2Json ()
		{
#if !UNITY_WEBPLAYER
			Object[] obj = Selection.objects;

			List<string> filePathList = new List<string>();

			for(int i=0; i<obj.Length; i++)
			{
				string path = AssetDatabase.GetAssetPath(obj[i]);
				path = Path.Combine(Directory.GetCurrentDirectory(), path);

				for(int j=0; j<_readExtention.Length; j++)
				{
					if(path.Contains(_readExtention[j]))
					{
						if(!filePathList.Contains(path))
							filePathList.Add(path);
					}
				}
			}

			for (int i = 0; i < filePathList.Count; i++) 
			{
				AddExcelFile (filePathList [i]);
			}

			EditorWindow.GetWindow<ExcelToJsonEditor>("Excels2JsonEditor");
#else
			CustomEditorTools.DisplayNoticeDialog ("Error", "The editor does not work on the Web Player environment. We recommend a PC environment.");
			return;
#endif
		}

		void OnGUI()
		{
			DisplayOutputJson();
			DisplayLoadExcelFolder();
			DisplayExcelFile ();
		}

		void OnDestroy()
		{
			if (_currentBookInfo != null) {
				_currentBookInfo.Clear ();
			}
		}

		void DisplayOutputJson()
		{
			EditorGUILayout.Separator();
			CustomEditorTools.DrawHeader("Output Excel To Json");
			CustomEditorTools.BeginContents();
			EditorGUILayout.BeginHorizontal();

			string outputPath = PlayerPrefs.GetString ("ExcelToJson_OutputPath");
			if (string.IsNullOrEmpty (outputPath)) 
			{
				outputPath = Application.dataPath + "/" + _jsonFolder;
				PlayerPrefs.SetString ("ExcelToJson_OutputPath", outputPath);
			}

			DirectoryInfo dirInfo = new DirectoryInfo (outputPath);
			if (!dirInfo.Exists) 
			{
				Directory.CreateDirectory (outputPath);
			}

			EditorGUILayout.TextField(outputPath);
			if(CustomEditorTools.DrawMiniButton("Choose"))
			{
				string changePath = EditorUtility.OpenFolderPanel("Output Json Folder", outputPath, "");

				if (!string.IsNullOrEmpty (changePath))
				{
					outputPath = changePath;
					PlayerPrefs.SetString ("ExcelToJson_OutputPath", outputPath);
				}
			}
			EditorGUILayout.EndHorizontal();
			CustomEditorTools.EndContens();
		}

		void DisplayLoadExcelFolder()
		{
			EditorGUILayout.Separator();
			CustomEditorTools.DrawHeader("Load Excel Folder");
			CustomEditorTools.BeginContents();
			EditorGUILayout.BeginHorizontal();

			string loadPath = PlayerPrefs.GetString ("ExcelLoadPath");
			if (string.IsNullOrEmpty (loadPath)) 
			{
				loadPath = Application.dataPath + "/" + _excelFolder;
				PlayerPrefs.SetString ("ExcelLoadPath", loadPath);
			}

			DirectoryInfo dirInfo = new DirectoryInfo (loadPath);
			if (!dirInfo.Exists) 
			{
				Directory.CreateDirectory (loadPath);
			}

			EditorGUILayout.TextField(loadPath);
			if(CustomEditorTools.DrawMiniButton("Choose"))
			{
				string changePath = EditorUtility.OpenFolderPanel("Output Json Folder", loadPath, "");

				if (!string.IsNullOrEmpty (changePath))
				{
					loadPath = changePath;
					PlayerPrefs.SetString ("ExcelLoadPath", loadPath);
				}
			}
			if(CustomEditorTools.DrawMiniButton("Load"))
			{
				DirectoryInfo loadDir = new DirectoryInfo(loadPath);
				if (!loadDir.Exists)
				{
					CustomEditorTools.DisplayNoticeDialog ("Warning", "Not Exist Excel Directory");
					return;
				}

				_currentBookInfo.Clear ();
				FileInfo[] file = loadDir.GetFiles();
				for(int i=0; i<file.Length; i++)
				{
					string[] extention = file[i].Name.Split('.');
					int lastIndex	   = extention.Length -1;
					bool isRead		   = false;
					for(int j=0; j<_readExtention.Length; j++)
					{
						if(extention[lastIndex].Equals(_readExtention[j]))
						{
							isRead = true;
							break;
						}
					}

					if (isRead) 
					{
						AddExcelFile (loadPath + "/" + file [i].Name);
					}
				}

				if (_currentBookInfo.Count <= 0) 
				{
					CustomEditorTools.DisplayNoticeDialog ("Warning", "Not Exist Excel File");
					return;
				}
			}
			EditorGUILayout.EndHorizontal();
			CustomEditorTools.EndContens();
		}

		static void AddExcelFile(string filePath)
		{
			if (IsContainExcelFile (filePath)) {
				CustomEditorTools.DisplayNoticeDialog ("Warning", "Already Add Excel File");
				return;
			}

			Book book = new Book();
			book.filePath = filePath;
			book.SetBook();

			if(book.isData)
			{
				_currentBookInfo.Add(book);
			}
		}

		static bool IsContainExcelFile(string filePath)
		{
			for (int i = 0; i < _currentBookInfo.Count; i++)
			{
				if (_currentBookInfo [i] != null) 
				{
					if (_currentBookInfo [i].fileName.Equals (Book.GetFileName(filePath)))
						return true;
				}
			}

			return false;
		}

		void DisplayExcelFile()
		{
			string outputPath = PlayerPrefs.GetString ("ExcelToJson_OutputPath");
			CustomEditorTools.DrawHeader("Excel File Info");
			CustomEditorTools.BeginContents();
			EditorGUILayout.BeginHorizontal();
			if (CustomEditorTools.DrawMiniButton ("Clear Excel Data"))
			{
				_currentBookInfo.Clear ();
			}
			if (CustomEditorTools.DrawMiniButton ("Remove All Json"))
			{
				DirectoryInfo outputDir = new DirectoryInfo (outputPath);
				if (outputDir.Exists)
				{
					FileInfo[] files = outputDir.GetFiles ();
					for (int i = files.Length - 1; i >= 0; i--)
					{
						FileUtil.DeleteFileOrDirectory (outputPath + "/" + files[i].Name);
					}
				}
				AssetDatabase.Refresh ();
			}
			if (CustomEditorTools.DrawMiniButton ("All Convert"))
			{
				if (_currentBookInfo.Count <= 0) 
				{
					CustomEditorTools.DisplayNoticeDialog ("Warring", "Excel Data Empty, Select Excel Data");
					return;
				}
				
				for (int i = 0; i < _currentBookInfo.Count; i++)
				{
					for (int j = 0; j < _currentBookInfo [i].sheet.Length; j++) 
					{
						ExcelToJsonConvert.ConvertToJson(_currentBookInfo [i].sheet[j], outputPath);
					}
				}

				CustomEditorTools.DisplayNoticeDialog ("Notice", "Completed Convert Excel To Json");
			}
			EditorGUILayout.EndHorizontal();
			CustomEditorTools.EndContens();
			CustomEditorTools.BeginContents();
			_scrollPos = EditorGUILayout.BeginScrollView(_scrollPos,GUILayout.MinHeight(200f));
			for(int i=_currentBookInfo.Count-1; i>= 0; i--)
			{
				CustomEditorTools.BeginContents();
				EditorGUILayout.BeginHorizontal();
				bool foldOut = GUILayout.Toggle(_currentBookInfo[i].foldOut,  _currentBookInfo[i].fileName, "Foldout", GUILayout.MaxWidth(200f), GUILayout.MinWidth(200f));

				if(foldOut != _currentBookInfo[i].foldOut)
				{
					_currentBookInfo[i].foldOut= foldOut;
				}
				EditorGUILayout.EndHorizontal();

				if(_currentBookInfo[i].foldOut)
				{
					if(_currentBookInfo[i].sheet == null) continue;

					for(int j=0; j<_currentBookInfo[i].sheet.Length; j++)
					{
						DisplaySheetInfo (_currentBookInfo [i].sheet [j]);
					}
				}
				CustomEditorTools.EndContens();
			}
			EditorGUILayout.EndScrollView();
			CustomEditorTools.EndContens();
		}

		void DisplaySheetInfo(Sheet sheet)
		{
			string outputPath = PlayerPrefs.GetString ("ExcelToJson_OutputPath");
			EditorGUILayout.BeginHorizontal(GUILayout.MinHeight(30));

			GUILayout.Space(10f);
			CustomEditorTools.DrawLabel(sheet.sheetName, 100f, "OL Title");
			GUILayout.Space(5f);
			if(CustomEditorTools.DrawMiniButton("Convert"))
			{
				ExcelToJsonConvert.ConvertToJson(sheet, outputPath);
				CustomEditorTools.DisplayNoticeDialog ("Notice", "Completed Convert Excel To Json");
				return;
			}

			GUILayout.Space(5f);
			EditorGUILayout.LabelField(sheet.outputName + ".json");
			EditorGUILayout.EndHorizontal();
		}
	}
}
