using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections.ObjectModel;

namespace Excel_Flat_File_Converter {
	class Program {

		// E:\Documents\Design\Other\Excel Flat File Converter\Test Files\Valid Test File.xlsx

		static StreamWriter sw;

		static List<string> files;
		static bool silent = false;
		static string[] validExtensions = new string[] { ".xlsx", ".xlsm", ".xls" };
		static string suffix = " (FLAT)";

		static void Main (string[] args) {
			files = new List<string> ();
			if (args.Length == 0) {
				GetFilesFromInput ();
			} else {
				for (int i = 0; i < args.Length; i++) {
					if (args[i].Substring(0, 1) == "-") {
						if (args[i].Substring(1) == "s") {
							// Silent mode
							silent = true;
						}
					}
				}
			}


			if (!silent) {
				Print ("Press any key to exit...");
				Console.ReadKey ();
			}

			sw.Close ();
		}

		static void GetFilesFromInput() {
			while (true) {
				Print ("Please enter the path to a file/folder to be converted:");
				string input = Console.ReadLine ();

				if (input == "exit" || input == "quit") {
					Environment.Exit (0);
				}

				if (File.Exists (input)) {
					if (validExtensions.Contains(Path.GetExtension (input))) {
						string filename = Path.GetFileNameWithoutExtension (input);
						if (filename.Length >= 6 && filename.Substring (filename.Length - suffix.Length, suffix.Length) != suffix) {
							files.Add (input);
							Print ("Success.", true);
							break;
						} else {
							Print ("ERROR: This file has already been converted.");
						}
					} else {
						Print ("ERROR: File is not an Excel workbook file. (.xlsx, .xlsm, .xls)");
					}
				} else {
					Print ("ERROR: File does not exist.");
				}
			}
		}

		static void Print(string msg, bool saveToLog = false) {
			// Create new StreamWriter if it doesn't exist yet
			if (sw == null) {
				sw = File.AppendText (Path.Combine (AppDomain.CurrentDomain.BaseDirectory, "history.log"));
			}

			// Write to Console and Log file
			if (!silent) Console.WriteLine (msg);
			if (saveToLog) sw.WriteLine ("[" + DateTime.Now.ToString () + "]  " + msg);
		}
	}
}
