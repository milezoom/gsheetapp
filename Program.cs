using System;
using System.Collections.Generic;
using Gtk;

namespace GSheetApp
{
	class MainClass
	{
		private static string Host = "10.9.8.153";
		private static string DBName = "bi";
		private static string Username = "userBI";
		private static string Password = "userBI";
		private static List<List<string>> queryResult = new List<List<string>> ();

		public static void Main (string[] args)
		{
			Application.Init ();
			MainWindow win = new MainWindow ();
			win.Show ();
			Application.Run ();
		}

		public static int ExecQuery (string query)
		{
			Database db = new Database (Host, DBName, Username, Password);
			queryResult = db.runQuery (query);
			return queryResult.Count;
		}

		public static int[] WriteToSheet (string filename)
		{
			SpreadsheetInteraction sheetInteraction = new SpreadsheetInteraction ();
			return sheetInteraction.writeToSheet (queryResult, filename);
		}

		public static string GetEmail ()
		{
			SpreadsheetInteraction sheetInteraction = new SpreadsheetInteraction ();
			return sheetInteraction.getEmail ();
		}

		public static void Reauthorize ()
		{
			SpreadsheetInteraction sheetInteraction = new SpreadsheetInteraction ();
			sheetInteraction.reauthorize ();
		}
	}
}
