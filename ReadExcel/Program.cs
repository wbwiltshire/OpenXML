using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


//
// Links: https://docs.microsoft.com/en-us/office/open-xml/
//		  https://docs.microsoft.com/en-us/office/open-xml/spreadsheets
//		  https://docs.microsoft.com/en-us/office/open-xml/how-to-open-a-spreadsheet-document-for-read-only-access
//        https://docs.microsoft.com/en-us/office/open-xml/how-to-parse-and-read-a-large-spreadsheet
//		  https://docs.microsoft.com/en-us/office/open-xml/how-to-retrieve-the-values-of-cells-in-a-spreadsheet
//

namespace ReadExcel
{
	class Program
	{
		static void Main(string[] args)
		{
			SheetData sheetData = null;
			SharedStringItem item = null;
			WorkbookPart workbookPart = null;
			Workbook workbook = null;
			WorksheetPart worksheetPart = null;
			string filePath = @"C:\Source\CSharp\Office\OpenXML\Sample.xlsx";
			StringBuilder text = new StringBuilder();
			string sheetId = String.Empty;
			int id = 0;

			try
			{
				// Open a SpreadsheetDocument for read-only access based on a filepath.
				using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
				{
					workbookPart = spreadsheetDocument.WorkbookPart;
					workbook = workbookPart.Workbook;

					Console.WriteLine("Found the following worksheets: ");
					foreach (var sheet in workbook.Sheets)
					{
						Console.WriteLine($"\t{((Sheet)sheet).Name}");
						if (((Sheet)sheet).Name == "Sheet1")
							sheetId = ((Sheet)sheet).Id;
					}

					// Print Sheet 1
					Console.WriteLine("\nContents of Sheet1:");
					worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
					sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
					foreach (Row r in sheetData.Elements<Row>())
					{
						text.Append("\t|");
						foreach (Cell c in r.Elements<Cell>())
						{
							if (c.DataType != null)
							{
								switch (c.DataType.Value)
								{
									case CellValues.Boolean:
										text.Append(Convert.ToInt32(c.InnerText)).Append("|");
										break;
									case CellValues.SharedString:
										if (Int32.TryParse(c.InnerText, out id))
										{
											item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
											if (item.Text != null)
											{
												text.Append(item.Text.Text).Append("|");
											}
											else if (item.InnerText != null)
											{
												text.Append(item.InnerText).Append("|");
											}
											else if (item.InnerXml != null)
											{
												text.Append(item.InnerXml).Append("|");
											}
										}
										break;
									case CellValues.String:
										Console.WriteLine("It's a string");
										break;
									case CellValues.Date:
										Console.WriteLine("It's a date");
										break;
									case CellValues.Number:
										Console.WriteLine("It's a number");
										break;
								}
							}
							else
							{
								text.Append(Convert.ToDecimal(c.InnerText)).Append("|");
							}
						}
						Console.WriteLine($"{text}");
						text.Clear();
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Exception: {ex.Message}");
			}
			
			Console.WriteLine("Press <enter> to end....");
			Console.ReadLine();
		}
	}
}
