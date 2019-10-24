using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace OpenSDKTools.Excel
{
	class DocumentWriter
	{
		internal void Write(string file, Dictionary<string, List<Variable>> variables, Dictionary<string, List<Table>> tables, List<Chart> charts)
		{
			using (var document = SpreadsheetDocument.Open(file, true))
			{
				IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

				foreach(var sheet in sheets)
				{
					var worksheetPart = GetWorksheetPart(document, sheet);
					var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

					if (variables.ContainsKey(sheet.Name))
					{
						foreach (var variable in variables[sheet.Name])
						{
							var cell = GetCell(worksheetPart.Worksheet, variable.RowIndex, variable.CellReference);
							cell.CellValue = new CellValue(variable.Value);
						}
					}

					if (tables.ContainsKey(sheet.Name))
					{
						foreach (var table in tables[sheet.Name])
						{
							var chart = charts.FirstOrDefault(x => x.TableKey == table.Key);

							var rowIndex = table.RowIndex;
							var rowNumber = 0;

							foreach (var row in table.Rows)
							{
								var wRow = new Row
								{
									RowIndex = rowIndex,
									Spans = new ListValue<StringValue>() { InnerText = $"1:{row.Cells.Count}" }
								};

								var column = table.ColumnName;
								var lastColumn = column;
								foreach (var cell in row.Cells)
								{
									var wCell = new Cell
									{
										CellReference = $"{column}{rowIndex}",
										CellValue = new CellValue(cell.Value)
									};

									wRow.Append(wCell);
									lastColumn = column;
									column = IncColumn(column);
								}

								if (chart != null)
								{
									if (rowNumber == chart.AxisXRowNumber)
									{
										chart.AxisX = $"'{sheet.Name}'!${table.ColumnName}${rowIndex}:${lastColumn}${rowIndex}";
									}

									if (rowNumber == chart.AxisYRowNumber)
									{
										chart.AxisY = $"'{sheet.Name}'!${table.ColumnName}${rowIndex}:${lastColumn}${rowIndex}";
									}
								}

								sheetData.Append(wRow);
								rowNumber++;
								rowIndex++;
							}
						}
					}
				}

				var chartWriter = new ChartWriter();

				foreach(var chart in charts)
				{
					var sheet = sheets.FirstOrDefault(x => x.Name == chart.SheetName);
					if (sheet != null)
					{
						var worksheetPart = GetWorksheetPart(document, sheet);

						chartWriter.Write(worksheetPart, chart);
					}
				}
			}
		}

		public static WorksheetPart GetWorksheetPart(SpreadsheetDocument document, string sheetName)
		{
			var sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().FirstOrDefault(x => x.Name == sheetName);

			if(sheet != null)
			{
				return GetWorksheetPart(document, sheet);
			}

			return null;
		}

		private static WorksheetPart GetWorksheetPart(SpreadsheetDocument document, Sheet sheet)
		{
			string relationshipId = sheet.Id.Value;
			WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

			return worksheetPart;
		}

		private static Cell GetCell(Worksheet worksheet, uint rowIndex, string cellReference)
		{
			Row row = GetRow(worksheet, rowIndex);

			if (row == null)
				return null;

			return row.Elements<Cell>().Where(c => string.Compare
				   (c.CellReference.Value, cellReference +
				   rowIndex, true) == 0).First();
		}

		private static Row GetRow(Worksheet worksheet, uint rowIndex)
		{
			return worksheet.GetFirstChild<SheetData>().
			  Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
		}

		private static string IncColumn(string column)
		{
			if(column.Length == 0)
			{
				return "A";
			}

			column = column.ToUpper();
			var lastChar = column.Last();
			var nextChar = NextLetter(lastChar);
			if (nextChar != 'A')
			{
				return column.Remove(column.Length - 1, 1) + nextChar;
			}
			else
			{
				return IncColumn(column.Remove(column.Length - 1, 1)) + nextChar;
			}
		}

		private static char NextLetter(char letter)
		{
			if (letter == 'Z')
			{
				return 'A';
			}
			else
			{
				return (char)(((int)letter) + 1);
			}
		}
	}
}
