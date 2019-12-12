/**
 * @brief  Define the following enum values in Excel, you can create a user-specific Json.
 * 		   Note the Excel / CustomExcel
 */

namespace ExcelToJson
{
	public enum JSON_TYPE
	{
		ERROR,
		JSON_OUTPUT,				/**< json output name																   */
		JSON_VERSION,				/**< json version code																   */
		JSON_START,					/**< json object start 																   */
		JSON_END,					/**< json object end 					 											   */
		JSON_ARRAY_START,			/**< json array start (if you use this, It should put the name of the array)		   */
		JSON_ARRAY_END,				/**< json array end 					 											   */
		JSON_TABLE_START,			/**< json table start (if you use this, It should put the name of the table)		   */
		JSON_TABLE_FIELD,			/**< you can not use in excel, only code 											   */
		JSON_TABLE_VALUE,			/**< you can not use in excel, only code 											   */
		JSON_TABLE_END,				/**< json table end																	   */
		JSON_VALUE,					/**< json value (if you use this, you can string, int, float, bool, null)		       */
	}
}