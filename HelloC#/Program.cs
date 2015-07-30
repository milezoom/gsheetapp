using System;
using System.IO;
using System.Threading;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Util.Store;
using Google.Apis.Services;

using Google.GData.Client;
using Google.GData.Spreadsheets;

using MySql.Data.MySqlClient;

namespace HelloC
{
	class MainClass
	{
		static string[] Scope = {
			DriveService.Scope.Drive + " " +
			"https://spreadsheets.google.com/feeds"
		};
		static string ApplicationName = "GSheetApp";
		static string RedirectUri = "urn:ietf:wg:oauth:2.0:oob";

		public static void Main (string[] args)
		{
			Console.WriteLine ("Application Started.");

			UserCredential credential;
			var stream = new FileStream ("client_secret.json", FileMode.Open, FileAccess.Read);
			ClientSecrets secret = GoogleClientSecrets.Load (stream).Secrets;

			using (stream) {
				string credPath = Directory.GetCurrentDirectory ();
				credPath = Path.Combine (credPath, ".credentials");

				credential = GoogleWebAuthorizationBroker.AuthorizeAsync (
					secret,
					Scope,
					"user",
					CancellationToken.None,
					new FileDataStore (credPath, true)
				).Result;
			}

			var driveService = new DriveService (new BaseClientService.Initializer () {
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});

			Console.WriteLine ();
			Console.WriteLine ("Here are your GSheet files in your Drive :");
			Console.WriteLine ();
			var files = driveService.Files.List ().Execute ().Items;
			foreach (Google.Apis.Drive.v2.Data.File file in files) {
				if (!(bool)file.ExplicitlyTrashed && file.MimeType == "application/vnd.google-apps.spreadsheet") {
					Console.WriteLine (file.Title + " ( " + file.Id + " )");
				}
			}

			OAuth2Parameters parameters = new OAuth2Parameters ();
			parameters.ClientId = secret.ClientId;
			parameters.ClientSecret = secret.ClientSecret;
			parameters.RedirectUri = RedirectUri;
			parameters.Scope = Scope.ToString ();
			parameters.AccessToken = credential.Token.AccessToken;

			GOAuth2RequestFactory requestFactory =
				new GOAuth2RequestFactory (null, "GSheetApp-v1", parameters);
			SpreadsheetsService ssService = new SpreadsheetsService ("GSheetApp-v1");
			ssService.RequestFactory = requestFactory;
			SpreadsheetQuery ssQuery = new SpreadsheetQuery ();
			SpreadsheetFeed ssFeed = ssService.Query (ssQuery);
			SpreadsheetEntry spreadsheet = null;
			for (int i = 0; i < ssFeed.TotalResults; i++) {
				if (ssFeed.Entries [i].Title.Text == "Hello From C#") {
					spreadsheet = (SpreadsheetEntry)ssFeed.Entries [i];
				}
			}
			WorksheetFeed wsFeed = spreadsheet.Worksheets;
			WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries [0];

			CellQuery cellQuery = new CellQuery (worksheet.CellFeedLink);
			CellFeed cellFeed = ssService.Query (cellQuery);

			/*------------- SQL Connection ---------------*/
			MySqlConnection connection;
			string connectionString = null;
			string query = null;

			connectionString = "SERVER=10.9.8.153;" + "DATABASE=bi;" +	"UID=userBI;" +	"PASSWORD=userBI;";
			connection = new MySqlConnection (connectionString);
			query = "select * from ot limit 5";

			try {
				connection.Open ();
				MySqlCommand command = new MySqlCommand (query, connection);
				MySqlDataReader reader = command.ExecuteReader ();
				uint counter = 1;
				CellEntry cellEntry;
				while (reader.Read ()) {
					cellEntry = new CellEntry (counter, 1, reader ["order_number"].ToString ());
					cellFeed.Insert (cellEntry);
					cellEntry = new CellEntry (counter, 2, reader ["customer_name"].ToString ());
					cellFeed.Insert (cellEntry);
					counter++;
				}
				connection.Close ();	
			} catch (Exception ex) {
				Console.WriteLine (ex);
			}
		}
	}
}
