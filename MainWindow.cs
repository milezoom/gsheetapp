using System;
using Gtk;
using GSheetApp;

public partial class MainWindow: Gtk.Window
{
	public MainWindow () : base (Gtk.WindowType.Toplevel)
	{
		Build ();
	}

	protected void OnDeleteEvent (object sender, DeleteEventArgs a)
	{
		Application.Quit ();
		a.RetVal = true;
	}

	protected void ExecQuery (object sender, EventArgs e)
	{
		string query = inpQuery.Text;
		int result = MainClass.ExecQuery (query);
		if (result == 0) {
			lbQueryStatus.Text = "Bad Query/No Result";
			inpQuery.Text = "";
		} else {
			lbQueryStatus.Text = "Success";
			lbQueryResult.Text = result + " Row(s)";
			inpQuery.Text = "";
		}
	}

	protected void UpdateQueryStatus (object sender, EventArgs e)
	{
		lbQueryStatus.Text = "Executing Query ....";
	}

	protected void WriteToSheet (object sender, EventArgs e)
	{
		string filename = inpSheetName.Text;
		int[] result = MainClass.WriteToSheet (filename);
		if (result [0] == 0) {
			lbWritingStatus.Text = "Write Failed/No Data From Input";
			inpSheetName.Text = "";
		} else {
			lbWritingStatus.Text = "Success";
			inpSheetName.Text = result [0] + " Rows and " + result [1] + " Cols";
			inpSheetName.Text = "";
		}
	}

	protected void UpdateWritingStatus (object sender, EventArgs e)
	{
		lbWritingStatus.Text = "Writing To Sheet ....";
	}
}
