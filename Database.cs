using System;
using MySql.Data.MySqlClient;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace GSheetApp
{
	public class Database
	{
		private MySqlConnection connection;

		public Database (string host, string dbName, string username, string password)
		{
			string connectionString = 
				"SERVER=" + host + ";" +
				"DATABASE=" + dbName + ";" +
				"UID=" + username + ";" +
				"PASSWORD=" + password + ";" +
				"Allow Zero Datetime=TRUE" + ";";

			this.connection = new MySqlConnection (connectionString);
		}

		public List<List<string>> runQuery (string query)
		{
			string cleanQuery = sanitizeQuery (query);
			var result = new List<List<string>> ();
			try {
				this.connection.Open ();

				MySqlCommand command = new MySqlCommand (cleanQuery, this.connection);
				MySqlDataReader reader = command.ExecuteReader ();

				while (reader.Read ()) {
					List<string> data = new List<string> ();
					for (int i = 0; i < reader.FieldCount; i++) {
						data.Add (reader [i].ToString ());
					}
					result.Add (data);
				}

				this.connection.Close ();
			} catch (Exception e) {
				Console.WriteLine (e);
			}
			return result;
		}

		public string sanitizeQuery (string query)
		{
			string result = "";

			// clean malicious command
			var regexAlter = new Regex (";alter ", RegexOptions.IgnoreCase);
			var regexUpdate = new Regex (";update ", RegexOptions.IgnoreCase);
			var regexDelete = new Regex (";delete ", RegexOptions.IgnoreCase);
			var regexInsert = new Regex (";insert ", RegexOptions.IgnoreCase);
			var regexDrop = new Regex (";drop ", RegexOptions.IgnoreCase);

			result = regexAlter.Replace (query, "");
			result = regexDelete.Replace (result, "");
			result = regexDrop.Replace (result, "");
			result = regexInsert.Replace (result, "");
			result = regexUpdate.Replace (result, "");
			result = Regex.Replace (result, "([Aa][Ll][Tt][Ee][Rr])", "");
			result = Regex.Replace (result, "([Dd][Ee][Ll][Ee][Tt][Ee])", "");
			result = Regex.Replace (result, "([Dd][Rr][Oo][Pp])", "");
			result = Regex.Replace (result, "([Ii][Nn][Ss][Ee][Rr][Tt])", "");
			result = Regex.Replace (result, "([Uu][Pp][Dd][Aa][Tt][Ee])", "");

			// clean malicious char
			var regexComment = new Regex (";--");
			result = regexComment.Replace (query, "");
			result = Regex.Replace (result, "(--)", "");

			// clean additional terminate char
			result = Regex.Replace (result, ";", "");

			// add terminate char to end of query
			result = result + ";";

			return result;
		}
	}
}