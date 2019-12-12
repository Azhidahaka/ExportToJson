/**
 * @brief  Define the layout of the Editor used to ExcelToJsonEditor
 */

namespace ExcelToJson
{
	using UnityEditor;
	using UnityEngine;

	public class CustomEditorTools  {

		static public void DrawHeader(string text)
		{
			GUI.backgroundColor = new Color(0.8f, 0.8f, 0.8f);
			GUILayout.BeginHorizontal();
			GUILayout.Label(text, "OL Title", GUILayout.MinWidth(300f));
			GUILayout.EndHorizontal();
			GUI.backgroundColor = Color.white;
		}

		static public void DrawLabel(string text, float width, string type ="")
		{
			if(string.IsNullOrEmpty(type))
				GUILayout.Label(text, GUILayout.MinWidth(width), GUILayout.MaxWidth(width));
			else
				GUILayout.Label(text, type, GUILayout.MinWidth(width), GUILayout.MaxWidth(width));
		}

		static public void BeginContents(float space = 3.0f)
		{
			EditorGUILayout.BeginHorizontal(GUILayout.MinHeight(10f));
			GUILayout.BeginVertical("AS TextArea");
			GUILayout.Space(space);
		}

		static public void EndContens(float space = 3.0f)
		{
			GUILayout.Space(space);
			GUILayout.EndVertical();
			EditorGUILayout.EndHorizontal();
		}

		static public bool DrawMiniButton(string text, float width = 100f)
		{
			return GUILayout.Button(text, EditorStyles.miniButton, GUILayout.MaxWidth(width), GUILayout.MinWidth(width));
		}

		static public void DisplayNoticeDialog(string title, string text)
		{
			EditorUtility.DisplayDialog(title, text, "OK");
		}
	}
}