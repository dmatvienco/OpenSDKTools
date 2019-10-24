using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public enum BorderStyle
	{
		None = 0,
		Left = 1,
		Top = 2,
		Bottom = 4,
		Right = 8
	}

	[Flags]
	public enum FontStyle
	{
		Normal = 0,
		Bold = 1,
		Underline = 4,
		Italic = 8
	}

	public enum ParagraphStyle
	{
		Normal,
		Heading1,
		Heading2,
		Heading3,
		Heading4,
		Heading5
	}

	public enum TextAlignment
	{
		Left,
		Center,
		Right
	}

	public enum TextVerticalAlignment
	{
		Top,
		Center,
		Bottom
	}

	public enum WriterType
	{
		Image,
		Link,
		Table,
		Text,
		Paragraph,
		Variable,
		Chart
	}

	public enum BreakType
	{
		Column,
		Page,
		TextWrapping
	}

	public enum ParagraphSpacing
	{
		None,
		Space10pt
	}

	public enum PageBreak
	{
		None,
		Before
	}

	
}
