﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;
using System.Text.RegularExpressions;
using System.Net.Http;
using Utility;
using OfficeOpenXml;
using System.Reflection;
using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using KB.Processes;
using KB.Configuration;
using KB.Utility;
using System.Data;
using System.Data.OleDb;

namespace Utility
{
    public static class Excel
    {
        public static string PointerToCell(this string pointer, int defaultRow) => pointer.Any(char.IsDigit) ? pointer : pointer + defaultRow;

        public static string InsertValues(this string text, ExcelWorksheet worksheet, int defaultRow, bool recursive) =>
            InsertValues(text, worksheet, '#', defaultRow, recursive);
        public static string InsertValues(this string text, ExcelWorksheet worksheet, char separators, int defaultRow, bool recursive)
        {
            string[] pointers = text.GetPointers(separators);
            if (pointers.Length == 0) return text;
            foreach (string p in pointers.Distinct())
                try { text = text.Replace(separators + p + separators, worksheet.Cells[PointerToCell(p, defaultRow)].Text); }
                // In case that cell is empty or merged to another, remove pointer to avoid StackOverFlow exception;
                catch { text = text.Replace(separators + p + separators, ""); }
            return recursive ? InsertValues(text, worksheet, separators, defaultRow, recursive) : text;
        }

        /// <summary>
        /// Get array of pointers from string
        /// </summary>
        /// <param name="str">Input string</param>
        /// <param name="separator">char at the beginning and end of each pointer</param>
        /// <returns></returns>
        public static string[] GetPointers(this string str, char separators)
        {
            List<string> arr = new List<string>(str.Split('#')); // Split with separator #.
            for (int i = 1; i < arr.Count; i += 2) // Every second string.
                // Merge string to one before until is pointer.
                while (i < arr.Count && !(!arr[i].Contains(" ") && arr[i].All(c => char.IsLetter(c) || char.IsDigit(c)) && arr[i].Any(char.IsLetter) && arr[i].EndsWith(string.Concat(arr[i].Where(char.IsDigit)))))
                // Pointer must be without any space, contains only letters and digits, contains at least one letter, ends with his row number (no mixed chars and digits).
                {
                    arr[i - 1] += "#" + arr[i];
                    arr.RemoveAt(i);
                }
            // Any second string will be a pointer.
            return arr.Skip(1).Where((s, idx) => idx % 2 == 0).ToArray();
        }

        public static string[] GetSheets(string excelFilePath)
        {
            List<string> sheets = new List<string>();
            using (OleDbConnection connection =
                    new OleDbConnection((excelFilePath.TrimEnd().ToLower().EndsWith("x"))
                    ? "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + excelFilePath + "';" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
                    : "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + excelFilePath + "';Extended Properties=Excel 8.0;"))
            {
                connection.Open();
                DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow drSheet in dt.Rows)
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                    {
                        string s = drSheet["TABLE_NAME"].ToString();
                        sheets.Add(s.StartsWith("'") ? s.Substring(1, s.Length - 3) : s.Substring(0, s.Length - 1));
                    }
                connection.Close();
            }
            return sheets.ToArray();
        }
    }
}
