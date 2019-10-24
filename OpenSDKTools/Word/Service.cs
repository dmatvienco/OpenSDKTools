using System;
using System.Collections.Generic;
using System.Linq;
using OpenSDKTools.Word;
using DocumentFormat.OpenXml.Packaging;

namespace OpenSDKTools.Word
{
	public sealed class Service
	{
		private Dictionary<string, Container> Containers;
		private Dictionary<string, Variable> Variables;

		private System.IO.FileInfo file;

		public Service(System.IO.FileInfo file)
		{
			this.file = file;

			this.Containers = new Dictionary<string, Container>();
			this.Variables = new Dictionary<string, Variable>();
		}

		public void AddContainer(Container container)
		{
			if (this.Containers.ContainsKey(container.Key))
			{
				throw new AlreadyAddedException(container.Key);
			}

			this.Containers.Add(container.Key, container);
		}

		public void AddContainers(List<Container> list)
		{
			foreach (var k in list)
				this.AddContainer(k);
		}

		public void AddVariable(Variable variable)
		{
			if (this.Variables.ContainsKey(variable.Key))
			{
				throw new AlreadyAddedException(variable.Key);
			}

			this.Variables.Add(variable.Key, variable);
		}

		public void AddVariables(List<Variable> list)
		{
			foreach (var k in list)
				this.AddVariable(k);
		}

		public void Generate()
		{
			var dw = new OpenSDKTools.Word.DocumentWriter();
			dw.Write(this.file.FullName, this.Containers, this.Variables);
		}

		public static void UniqueVariablesCollection(string file1, string file2)
		{
			var names = new Dictionary<string, string>();
			using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file1, true))
			{
				var context = new Context()
				{
					Package = package
				};

				var mainPart = context.Package.MainDocumentPart;

				var markers = new List<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

				markers = mainPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>().ToList();

				// header
				if (mainPart.HeaderParts != null)
				{
					foreach (var part in mainPart.HeaderParts)
					{
						var _e = part.Header.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				// footer
				if (mainPart.FooterParts != null)
				{
					foreach (var part in mainPart.FooterParts)
					{
						var _e = part.Footer.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				// footnotes
				if (mainPart.FootnotesPart != null && mainPart.FootnotesPart.Footnotes != null)
				{
					foreach (var part in mainPart.FootnotesPart.Footnotes)
					{
						var _e = part.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				foreach (var marker in markers)
				{
					var tag = marker.SdtProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Tag>();
					if (tag != null && !names.ContainsKey(tag.Val.Value))
					{
						string _text = string.Empty;
						var sdtRun = marker.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.SdtContentRun>();
						if (sdtRun != null)
						{
							var run = sdtRun.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Run>();
							if (run != null)
							{
								var text = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>();
								if (text != null)
								{
									_text = text.Text;
								}
							}
						}

						names.Add(tag.Val.Value, _text);
					}
				}
			}

			using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file2, true))
			{
				var context = new Context()
				{
					Package = package
				};

				var mainPart = context.Package.MainDocumentPart;

				var markers = new List<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

				markers = mainPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>().ToList();

				// header
				if (mainPart.HeaderParts != null)
				{
					foreach (var part in mainPart.HeaderParts)
					{
						var _e = part.Header.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				// footer
				if (mainPart.FooterParts != null)
				{
					foreach (var part in mainPart.FooterParts)
					{
						var _e = part.Footer.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				// footnotes
				if (mainPart.FootnotesPart != null && mainPart.FootnotesPart.Footnotes != null)
				{
					foreach (var part in mainPart.FootnotesPart.Footnotes)
					{
						var _e = part.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				var names2 = new Dictionary<string, string>();
				foreach (var marker in markers)
				{
					var tag = marker.SdtProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Tag>();
					if (tag != null && !names.ContainsKey(tag.Val.Value) && !names2.ContainsKey(tag.Val.Value))
					{
						string _text = string.Empty;
						var sdtRun = marker.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.SdtContentRun>();
						if (sdtRun != null)
						{
							var run = sdtRun.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Run>();
							if (run != null)
							{
								var text = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>();
								if (text != null)
								{
									_text = text.Text;
								}
							}
						}

						names2.Add(tag.Val.Value, _text);
					}
				}

				names2.OrderBy(n => n.Key).ToList().ForEach(n => Console.WriteLine(string.Format("VariableKeys.Add(\"{0}\", \"{1}\");", n.Key, n.Value)));
			}
		}

		public static void PrepareVariablesCollection(string file)
		{
			using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, true))
			{
				var context = new Context()
				{
					Package = package
				};

				var mainPart = context.Package.MainDocumentPart;

				var markers = new List<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

				markers = mainPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>().ToList();

				// header
				if (mainPart.HeaderParts != null)
				{
					foreach (var part in mainPart.HeaderParts)
					{
						var _e = part.Header.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				// footer
				if (mainPart.FooterParts != null)
				{
					foreach (var part in mainPart.FooterParts)
					{
						var _e = part.Footer.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				// footnotes
				if (mainPart.FootnotesPart != null && mainPart.FootnotesPart.Footnotes != null)
				{
					foreach (var part in mainPart.FootnotesPart.Footnotes)
					{
						var _e = part.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				var names = new Dictionary<string, string>();
				foreach (var marker in markers)
				{
					var tag = marker.SdtProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Tag>();
					if (tag != null && !names.ContainsKey(tag.Val.Value))
					{
						string _text = string.Empty;
						var sdtRun = marker.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.SdtContentRun>();
						if (sdtRun != null)
						{
							var run = sdtRun.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Run>();
							if (run != null)
							{
								var text = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>();
								if (text != null)
								{
									_text = text.Text;
								}
							}
						}

						names.Add(tag.Val.Value, _text);
					}
				}

				names.OrderBy(n => n.Key).ToList().ForEach(n => Console.WriteLine(string.Format("VariableKeys.Add(\"{0}\", \"{1}\");", n.Key, n.Value)));
			}
		}

		public static void ReplaceMarkerTitlesWithNames(string file)
		{
			using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, true))
			{
				var context = new Context()
				{
					Package = package
				};

				var mainPart = context.Package.MainDocumentPart;

				var markers = new List<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

				markers = mainPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>().ToList();

				// header
				if (mainPart.HeaderParts != null)
				{
					foreach (var part in mainPart.HeaderParts)
					{
						var _e = part.Header.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				// footer
				if (mainPart.FooterParts != null)
				{
					foreach (var part in mainPart.FooterParts)
					{
						var _e = part.Footer.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				// footnotes
				if (mainPart.FootnotesPart != null && mainPart.FootnotesPart.Footnotes != null)
				{
					foreach (var part in mainPart.FootnotesPart.Footnotes)
					{
						var _e = part.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtElement>();

						markers.AddRange(_e);
					}
				}

				foreach(var marker in markers)
				{
					var tag = marker.SdtProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Tag>();
					if (tag != null)
					{
						string name = tag.Val.Value;
						var sdtRun = marker.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.SdtContentRun>();
						if (sdtRun != null)
						{
							var run = sdtRun.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Run>();
							if (run != null)
							{
								var text = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>();
								if (text != null)
								{
									text.Text = name;
								}
							}
						}
					}
				}
			}
		}

		public static uint GetMaxDocPrId(MainDocumentPart mainPart)
		{
			uint max = 1;

			//Get max id value of docPr elements
			foreach (var docPr in mainPart.Document.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>())
			{
				uint id = docPr.Id;
				if (id > max)
					max = id;
			}
			return max;
		}

		public static Container Parse(string content)
		{
			throw new NotImplementedException();
		}
	}
}
