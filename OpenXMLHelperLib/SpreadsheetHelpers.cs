using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLHelperLib
{
	public class SpreadsheetHelpers
	{
		public static ICollection<SheetData> GetSheetData(WorksheetPart worksheetPart)
		{
			ICollection<SheetData> sheetData = null;

			sheetData = worksheetPart.Worksheet.Elements<SheetData>().ToList();

			return sheetData;
		}

		public static ICollection<Row> GetRows(SheetData sheetData)
		{
			return sheetData.Elements<Row>().ToList();
		}

		public static ICollection<Cell> GetCells(Row row)
		{
			return row.Elements<Cell>().ToList();
		}

		public static SharedStringItem GetSharedStringTable(WorkbookPart workbookPart, int id)
		{
			SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);

			return item;
		}
	}
}
