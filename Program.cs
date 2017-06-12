using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excel_Flat_File_Converter {
	class Program {

		static StreamWriter sw;
		static Excel.Application xlApp;

		static List<string> files;
		static bool silent = false;

		static readonly string[] validExtensions = new string[] { ".xlsx", ".xlsm", ".xls" };
		static readonly string suffix = " (FLAT)";

		static void Main (string[] args) {

			files = new List<string> ();

			// Parse any given arguments
			if (args.Length == 0) {
				GetFilesFromInput ();		// If no arguments then ask the user for input
			} else {
				for (int i = 0; i < args.Length; i++) {
					if (args[i].Substring (0, 1) == "-") {
						if (args[i].Substring (1) == "s") {
							// Silent mode
							silent = true;
						}
					} else {
						if (File.Exists (args[i]) || Directory.Exists (args[i])) {
							GetFiles (args[i]);
						}
					}
				}
			}
			LogLine ();

			// Loop through each of the files and create a flat file
			LogLine ("Creating flat files...", true);
			if (files.Count > 0) {
				foreach (string file in files) {
					LogLine ("    " + Path.GetFileName (file), true);
					CreateFlatFile (file);
				}
			}
			LogLine ();

			// If we're not in silent mode, wait until the user presses a key to exit
			if (!silent) {
				LogLine ("Press any key to exit...");
				Console.ReadKey ();
			}

			// Finish
			sw.Close ();
			xlApp.Quit ();
		}


		static void LogLine (string msg = "", bool saveToLog = false) {
			// Create new StreamWriter if it doesn't exist yet
			if (sw == null) {
				sw = File.AppendText (Path.Combine (AppDomain.CurrentDomain.BaseDirectory, "history.log"));
			}

			// Write to Console and Log file
			if (!silent)
				Console.WriteLine (msg);
			if (saveToLog)
				sw.WriteLine ("[" + DateTime.Now.ToString () + "]  " + msg);
		}


		static bool GetFiles (string path) {
			if (File.Exists (path)) {
				if (validExtensions.Contains (Path.GetExtension (path))) {
					string filename = Path.GetFileNameWithoutExtension (path);
					if (filename.Length >= 6 && filename.Substring (filename.Length - suffix.Length, suffix.Length) != suffix) {
						files.Add (path);
						LogLine ("Adding file...", true);
						LogLine ("    " + path, true);
						return true;
					} else {
						LogLine ("ERROR: This file has already been converted.");
						return false;
					}
				} else {
					LogLine ("ERROR: File is not an Excel workbook file. (.xlsx, .xlsm, .xls)");
					return false;
				}
			} else {
				// add else if before this for Directory.Exists(...) check and take all valid files inside the folder
				LogLine ("ERROR: File does not exist.");
				return false;
			}
		}


		static void GetFilesFromInput () {
			while (true) {
				LogLine ("Please enter the path to a file/folder to be converted:");
				string input = Console.ReadLine ();

				if (input == "exit" || input == "quit") {
					Environment.Exit (0);
				}

				if (GetFiles (input)) {
					break;
				}

				LogLine ();
			}
		}

		
		static void CreateFlatFile (string file) {
			// Check file still exists
			if (!File.Exists(file)) {
				LogLine ("ERROR: Given file does not exist!", true);
				LogLine ("     : Cannot create a flat file version of this file.", true);
				return;
			}

			// Create new copy of the file
			string newfile = file.Substring (0, file.Length - Path.GetExtension (file).Length) + suffix + Path.GetExtension (file);
			if (File.Exists(newfile))
				LogLine ("Overwriting existing flat file...", true);
			try {
				File.Copy (file, newfile, true);
				File.SetLastWriteTime (newfile, DateTime.Now);
			} catch (IOException) {
				LogLine ("ERROR: Could not overwrite the existing flat file as it was in use by another program.", true);
				LogLine ("     : Please ensure it is closed in Excel and try again.", true);
				LogLine ();
			}

			// Create a new Excel Application instance if we haven't already
			if (xlApp == null) {
				xlApp = new Excel.Application ();
				xlApp.Visible = false;
				xlApp.AskToUpdateLinks = false;
			}

			// Open this file with the Excel Application
			Excel.Workbook wb = xlApp.Workbooks.Open (newfile, 0, false);
			foreach (Excel.Worksheet ws in wb.Sheets) {
				try {
					// Copy the entire sheet and paste as values
					ws.UsedRange.Copy (Type.Missing);
					ws.UsedRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);
				} catch (COMException e) {
					LogLine ("ERROR: An error occurred when pasting as values...", true);
					LogLine ("    Error Code: " + e.ErrorCode.ToString (), true);
					LogLine ("    " + e.Message, true);
				}
			}

			// Select cell A1 of the first sheet, just to reset all of the selections made when copying and pasting (tidier)
			Excel.Worksheet firstSheet = wb.Worksheets[1];
			firstSheet.Activate ();
			firstSheet.Range["ZZ999"].Activate ();		// Select something not selected to reset selection
			firstSheet.Range["A1"].Activate ();			// Select A1 on first sheet

			// Save and close!
			wb.Save ();
			wb.Close ();
		}
	}
}
