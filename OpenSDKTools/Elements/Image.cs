using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class Image : IElement, IWritable
	{
		public Image()
		{
			this.Key = string.Empty;
		}

		public string Id { get; set; }

		public string Key { get; set; }

		public byte[] Data { get; set; }

		public double Width { get; set; }

		public double Height { get; set; }

		IWriter IWritable.GetWriter()
		{
			return new Word.ImageWriter();
		}
	}
}
