using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace Excel_Flat_File_Converter {
	class Program {

		static StreamWriter sw;
		static Excel.Application xlApp;

		static List<string> files;
		static bool silent = false;
		static bool nomacros = false;
		static string suffix = " (FLAT)";

		static readonly string[] validExtensions = new string[] { ".xlsx", ".xlsm", ".xls" };

		static void Main (string[] args) {

			files = new List<string> ();

			LogLine ("Starting...", true);
			LogLine ();

			// Parse any given arguments
			if (args.Length == 0) {
				GetFilesFromInput ();		// If no arguments then ask the user for input
			} else {
				for (int i = 0; i < args.Length; i++) {
					if (args[i].Substring (0, 1) == "-" && args[i].Length > 1) {
						args[i] = args[i].ToLower ();
						if (args[i].Substring (1) == "s") {
							// Silent mode
							silent = true;
						}
						if (args[i].Substring (1) == "nomacros" || args[i].Substring (1) == "nm" || args[i].Substring (1) == "xlsx") {
							// Save as xlsx (to remove all macros embedded in the workbook
							nomacros = true;
						}
						if (args[i].Substring (1) == "datesuffix" || args[i].Substring (1) == "datesuffix1" || args[i].Substring (1) == "ds" || args[i].Substring (1) == "ds1") {
							// Set suffix to current date in the format " (YYYYMMDD)"
							suffix = " (" + DateTime.Now.ToString ("yyyyMMdd") + ")";
						}
						if (args[i].Substring (1) == "datesuffix2" || args[i].Substring (1) == "ds2") {
							// Set suffix to current date in the format " (YYYY_MM_DD)"
							suffix = " (" + DateTime.Now.ToString ("yyyy_MM_dd") + ")";
						}
						if (args[i].Substring (1) == "datesuffixreverse" || args[i].Substring (1) == "dsr" || args[i].Substring (1) == "dsr1") {
							// Set suffix to current date in the format " (DD_MM_YYYY)"
							suffix = " (" + DateTime.Now.ToString ("dd_MM_yyyy") + ")";
						}
					} else {
						if (File.Exists (args[i]) || Directory.Exists (args[i])) {
							GetFiles (args[i]);
						} else {
							LogLine ("The following file or folder given as an argument does not exist...", true);
							LogLine ("    " + args[i], true);
						}
					}
				}
			}
			LogLine ();

			// Loop through each of the files and create a flat file
			if (files.Count > 0) {
				LogLine ("Creating flat files...", true);
				foreach (string file in files) {
					LogLine ("    " + Path.GetFileName (file), true);
					CreateFlatFile (file);
				}
			} else {
				LogLine ("No files given.", true);
			}
			LogLine ();

			LogLine ("Complete!" + Environment.NewLine, true);

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
						LogLine ("Added file:", true);
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
				LogLine ("        ERROR: Given file does not exist!", true);
				LogLine ("             : Cannot create a flat file version of this file.", true);
				return;
			}

			// Create a new Excel Application instance if we haven't already
			if (xlApp == null) {
				LogLine ("        Opening Excel...", true);
				xlApp = new Excel.Application ();
				xlApp.Visible = false;
				xlApp.AskToUpdateLinks = false;
				xlApp.DisplayAlerts = false;
			}

			// Create new copy of the file
			string newfile = file.Substring (0, file.Length - Path.GetExtension (file).Length) + suffix + Path.GetExtension (file);
			if (File.Exists(newfile))
				LogLine ("        Overwriting existing flat file...", true);
			try {
				File.Copy (file, newfile, true);
				File.SetLastWriteTime (newfile, DateTime.Now);
			} catch (IOException) {
				LogLine ("        ERROR: Could not overwrite the existing flat file as it was in use by another program.", true);
				LogLine ("             : Please ensure it is closed in Excel and try again.", true);
				LogLine ();
			}

			// Open this file with the Excel Application
			Excel.Workbook wb = xlApp.Workbooks.Open (newfile, 0, false);
			foreach (Excel.Worksheet ws in wb.Sheets) {
				try {
					// Copy the entire sheet and paste as values
					ws.UsedRange.Copy (Type.Missing);
					ws.UsedRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);

					// Reset selection to cell A1 for this sheet (tidier than leaving UsedRange selected)
					ws.Activate ();
					ws.Range["ZZ99"].Activate ();
					ws.Range["A1"].Activate ();
				} catch (COMException e) {
					LogLine ("        ERROR: An COMException occurred when pasting as values...", true);
					LogLine ("            Error Code: " + e.ErrorCode.ToString (), true);
					LogLine ("            " + e.Message, true);
				}
			}

			// Go back to first sheet once we've converted the file
			(wb.Worksheets[1] as Excel.Worksheet).Activate ();
			(wb.Worksheets[1] as Excel.Worksheet).Range["A1"].Activate ();

			// Save and close!
			if (nomacros && Path.GetExtension (file) == ".xlsm") {
				try {
					LogLine ("        Removing macros from the .xlsm file...", true);
					xlApp.DisplayAlerts = false;            // Make sure alerts are still turned off
					wb.SaveAs (Path.Combine (Path.GetDirectoryName (file), Path.GetFileNameWithoutExtension (file)) + suffix + ".xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook);
					// If we can successfully find this new .xlsx file then we can go ahead and delete the old temporary .xlsm version
					if (File.Exists (Path.Combine (Path.GetDirectoryName (file), Path.GetFileNameWithoutExtension (file)) + suffix + ".xlsx")) {
						File.Delete (Path.Combine (Path.GetDirectoryName (file), Path.GetFileNameWithoutExtension (file)) + suffix + ".xlsm");
					}
				} catch (COMException e) {
					LogLine ("        ERROR: An COMException occurred when saving the macro-free file...", true);
					LogLine ("            Error Code: " + e.ErrorCode.ToString (), true);
					LogLine ("            " + e.Message, true);
				} catch (IOException e) {
					LogLine ("        ERROR: An IOException occurred when saving the macro-free file...", true);
					LogLine ("            Error Code: " + e.HResult.ToString (), true);
					LogLine ("            " + e.Message, true);
				}
			} else {
				wb.Save ();
			}
			wb.Close ();
		}
	}
}
