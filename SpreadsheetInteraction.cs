using System;﻿
using System.IO;
using System.Threading;
using System.Collections.Generic;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Drive.v2;
using Google.Apis.Drive.v2.Data;
using Google.Apis.Plus.v1;
using Google.Apis.Util;
using Google.Apis.Util.Store;
using Google.Apis.Services;

using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace GSheetApp
{
	public class SpreadsheetInteraction
	{
		private string[] SCOPE = {
			DriveService.Scope.Drive + " " +
			"https://spreadsheets.google.com/feeds " +
			"https://www.googleapis.com/auth/userinfo.email"
		};
		private string APPLICATION_NAME = "GSheetApp";
		private string REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob";
		private string credentialPath = "";
		private UserCredential credential;
		private ClientSecrets secret;

		public SpreadsheetInteraction ()
		{
			Console.WriteLine ("Begin Create Spreadsheet Integration");
			var stream = new FileStream ("client_secret.json", FileMode.Open, FileAccess.Read);
			this.secret = GoogleClientSecrets.Load (stream).Secrets;
			Console.WriteLine ("Secret Saved");

			using (stream) {
				this.credentialPath = Path.Combine (Directory.GetCurrentDirectory (), ".credentials");

				this.credential = GoogleWebAuthorizationBroker.AuthorizeAsync (
					secret, SCOPE, "user", CancellationToken.None, new FileDataStore (credentialPath, true)
				).Result;
			}

			stream.Close ();
		}

		public DriveService createDriveService ()
		{
			return new DriveService (new BaseClientService.Initializer () {
				HttpClientInitializer = credential,
				ApplicationName = APPLICATION_NAME,
			});
		}

		public PlusService createPlusService ()
		{
			return new PlusService (new BaseClientService.Initializer () {
				HttpClientInitializer = credential,
				ApplicationName = APPLICATION_NAME,
			});
		}

		public string createSheet (DriveService service, string filename)
		{
			var file = new Google.Apis.Drive.v2.Data.File ();
			file.Title = filename;
			file.Description = string.Format ("Created via {0} at {1}", APPLICATION_NAME, DateTime.Now.ToString ());
			file.MimeType = "application/vnd.google-apps.spreadsheet";

			var request = service.Files.Insert (file);
			var result = request.Execute ();
			return result.Id;
		}

		public SpreadsheetsService createSheetService ()
		{
			OAuth2Parameters parameters = new OAuth2Parameters ();
			parameters.ClientId = this.secret.ClientId;
			parameters.ClientSecret = this.secret.ClientSecret;
			parameters.RedirectUri = REDIRECT_URI;
			parameters.Scope = SCOPE.ToString ();
			parameters.AccessToken = credential.Token.AccessToken;

			var spreadsheetsService = new SpreadsheetsService (APPLICATION_NAME);
			spreadsheetsService.RequestFactory = new GOAuth2RequestFactory (null, "GSheetApp-v1", parameters);
			return spreadsheetsService;
		}

		public SpreadsheetEntry getSheetEntry (SpreadsheetsService service, string fileId)
		{
			var query = new SpreadsheetQuery (
				            "https://spreadsheets.google.com/feeds/spreadsheets/" + fileId
			            );
			var feed = service.Query (query);
			return (SpreadsheetEntry)feed.Entries [0];
		}

		public WorksheetEntry getWorksheetEntry (SpreadsheetEntry entry, int wsId)
		{
			WorksheetFeed wsFeed = entry.Worksheets;
			return (WorksheetEntry)wsFeed.Entries [wsId];
		}

		public CellFeed getCellFeed (WorksheetEntry entry, SpreadsheetsService service)
		{
			CellQuery cellQuery = new CellQuery (entry.CellFeedLink);
			return service.Query (cellQuery);
		}

		public int[] writeToSheet (List<List<string>> data, string filename)
		{
			int[] result = { 0, 0 };
			string sheetId = createSheet (createDriveService (), filename);
			Console.WriteLine ("success creating new sheet");
			SpreadsheetEntry sheetEntry = getSheetEntry (createSheetService (), sheetId);
			Console.WriteLine ("success get sheet entry");
			WorksheetEntry wsEntry = getWorksheetEntry (sheetEntry, 0);
			wsEntry.Rows = Convert.ToUInt32 (data.Count);
			wsEntry.Cols = Convert.ToUInt32 (data [0].Count);
			wsEntry.Update ();
			Console.WriteLine ("success get worksheet entry");
			CellFeed feed = getCellFeed (wsEntry, createSheetService ());
			Console.WriteLine ("success get cell feed");
			CellEntry entry;

			try {
				for (uint i = 0; i < data.Count; i++) {
					List<string> row = data [Convert.ToInt32 (i)];
					for (uint j = 0; j < row.Count; j++) {
						entry = new CellEntry (i + 1, j + 1, (row [Convert.ToInt32 (j)]).ToString ());
						Console.WriteLine ("insert data into cel " + (i + 1) + (j + 1));
						feed.Insert (entry);
						Console.WriteLine ("success inserting data");
					}
				}
				Console.WriteLine ("success writing data to new sheet");
				result [0] = Convert.ToInt32 (wsEntry.Rows);
				result [1] = Convert.ToInt32 (wsEntry.Cols);
				return result;
			} catch (Exception e) {
				Console.WriteLine (e);
				return result;
			}
		}

		public string getEmail ()
		{
			string email = "Unknown";
			try {
				PlusService service = this.createPlusService ();
				PeopleResource.GetRequest people = service.People.Get ("me");
				var me = people.Execute ();
				email = me.Emails [0].Value.ToString ();	
			} catch (Exception e) {
				Console.WriteLine (e);
			}
			return email;
		}

		public void reauthorize ()
		{
			Array.ForEach (Directory.GetFiles (this.credentialPath),
				delegate(string path) {
					System.IO.File.Delete (path);
				}
			);
			try {
				this.createPlusService ();
			} catch (Exception e) {
				Console.WriteLine (e);
			}
		}
	}
}
