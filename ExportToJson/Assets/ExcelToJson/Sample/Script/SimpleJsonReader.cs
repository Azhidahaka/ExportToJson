/**
 * @brief A simple example of using the Json data reading
 * 		  When you read an Excel file using the default table type
 * 		  Check the file "Simple Data.xls"
 */

namespace ExcelToJson
{
	using UnityEngine;
	using System.Collections;

	public class SimpleJson
	{
		public SimpleJsonData[] SimpleData;
	}

	public class SimpleJsonData
	{
		public string strValue   = "";
		public int    intValue   = 0;
		public float  floatValue = 0;
	}

	public class SimpleJsonReader : MonoBehaviour {

		public TextAsset simpleData;
		SimpleJson json;

		void Awake()
		{
			if(simpleData == null) return;

			string str = simpleData.text;

			json	   = JsonFx.Json.JsonReader.Deserialize<ExcelToJson.SimpleJson>(str);

			Debug.Log("SimpleJsonData");
			if(json != null)
			{
				for(int i=0; i<json.SimpleData.Length; i++)
				{
					Debug.Log("DATA       : " + i.ToString() + " strValue   : " + json.SimpleData[i].strValue + 
					          " intValue   : " + json.SimpleData[i].intValue + " floatValue : " + json.SimpleData[i].floatValue);
				}
			}
		}
	}
}
