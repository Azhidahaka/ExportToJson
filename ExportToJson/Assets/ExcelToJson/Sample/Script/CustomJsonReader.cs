/**
 * @brief A Custom example of using the Json data reading
 * 		  When you read an Excel file using the Jsontype table 
 * 		  Check the file "Custom Data.xls"
 */

namespace ExcelToJson
{
	using UnityEngine;
	using System.Collections;

	public class CusomJson 
	{
		public CustomJsonData CustomData;
	}

	public class CustomJsonData
	{
		public int   		 ID;
		public int[] 		 ARRAY;
		public Child 		 ChildObject;
		public TableObject[] TableData;
	}

	public class Child
	{
		public string child1;
		public string child2;
	}

	public class TableObject
	{
		public int T1;
		public int T2;
		public int T3;
	}

	public class CustomJsonReader : MonoBehaviour {

		public TextAsset customData;
		CusomJson json;
		
		void Awake()
		{
			if(customData == null) return;

			string str = customData.text;
			
			json	   = JsonFx.Json.JsonReader.Deserialize<ExcelToJson.CusomJson>(str);
			DebugObject(json);
		}

		void DebugObject(ExcelToJson.CusomJson json)
		{
			Debug.Log("CustomData");
			Debug.Log("ID : " + json.CustomData.ID);
			for(int i=0; i<json.CustomData.ARRAY.Length; i++)
			{
				Debug.Log("ARRAY" +i.ToString() + " : " + json.CustomData.ARRAY[i]);
			}
			Debug.Log("ChildObject child1 : " + json.CustomData.ChildObject.child1);
			Debug.Log("ChildObject child2 : " + json.CustomData.ChildObject.child2);

			for(int i=0; i<json.CustomData.TableData.Length; i++)
			{
				Debug.Log("Table : " + i.ToString());
				Debug.Log("Table T1 : " + json.CustomData.TableData[i].T1);
				Debug.Log("Table T2 : " + json.CustomData.TableData[i].T2);
				Debug.Log("Table T3 : " + json.CustomData.TableData[i].T3);
			}
		}
	}
}