using System;
using System.Collections.Generic;
using System.Linq;
using OpenSDKTools.Word;
using DocumentFormat.OpenXml.Packaging;

namespace OpenSDKTools.Excel
{
	public sealed class Service
	{
		private Dictionary<string, List<Table>> Tables;
		private Dictionary<string, List<Variable>> Variables;
		private List<Chart> Charts;

		private System.IO.FileInfo file;

		public Service(System.IO.FileInfo file)
		{
			this.file = file;

			this.Tables = new Dictionary<string, List<Table>>();
			this.Variables = new Dictionary<string, List<Variable>>();
			this.Charts = new List<Chart>();
		}

		public void AddTable(string sheetName, Table table)
		{
			if (!this.Tables.ContainsKey(sheetName))
			{
				this.Tables.Add(sheetName, new List<Table>());
			}

			this.Tables[sheetName].Add(table);
		}

		public void AddTables(string sheetName, List<Table> list)
		{
			foreach (var k in list)
			{
				this.AddTable(sheetName, k);
			}
		}

		public void AddVariable(string sheetName, Variable variable)
		{
			if (!this.Variables.ContainsKey(sheetName))
			{
				this.Variables.Add(sheetName, new List<Variable>());
			}

			this.Variables[sheetName].Add(variable);
		}

		public void AddVariables(string sheetName, List<Variable> list)
		{
			foreach (var k in list)
			{
				this.AddVariable(sheetName, k);
			}
		}

		public void AddChart(Chart chart)
		{
			this.Charts.Add(chart);
		}

		public void AddCharts(List<Chart> list)
		{
			foreach (var k in list)
			{
				this.AddChart(k);
			}
		}

		public void Generate()
		{
			var dw = new OpenSDKTools.Excel.DocumentWriter();
			dw.Write(this.file.FullName, this.Variables, this.Tables, this.Charts);
		}
	}
}
