using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class List : IElement, IWritable
	{
		public string Key { get; set; }

		public List<Paragraph> Paragraphs { get; set; }

		public IWriter GetWriter()
		{
			throw new NotImplementedException();
		}
	}
}
