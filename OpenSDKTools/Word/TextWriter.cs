using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenSDKTools.Word
{
	class TextWriter : IWriter
	{
		public WriterType GetWriterType()
		{
			return WriterType.Text;
		}

		public void Write(SdtElement marker, Text text)
		{
			var parent = marker.Parent;

			var run = Generate(text);

			parent.InsertAfter(run, marker);
		}

		public static Run Generate(Text text)
		{
			var run = new Run();

			run.Append(new RunProperties());

			if (text.FontSize > 0)
			{
				run.RunProperties.Append(new FontSize()
				{
					Val = (text.FontSize * 2).ToString()
				});
			}

			AppendStyle(run.RunProperties, text.FontStyle);

			run.Append(new DocumentFormat.OpenXml.Wordprocessing.Text(text.Value) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });

			return run;
		}

		private static void AppendStyle(RunProperties properties, FontStyle style)
		{
			if ((style & FontStyle.Bold) == FontStyle.Bold)
			{
				properties.Append(
					new Bold()
				);
			}
			if ((style & FontStyle.Underline) == FontStyle.Underline)
			{
				properties.Append(
					new Underline() { Val = UnderlineValues.Single }
				);
			}

			if ((style & FontStyle.Italic) == FontStyle.Italic)
			{
				properties.Append(
					new Italic()
				);
			}
		}
	}
}
