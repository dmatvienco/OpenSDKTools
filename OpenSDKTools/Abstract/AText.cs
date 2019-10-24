using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public abstract class AText : IText
	{
		public FontStyle FontStyle { get; set; }

		public string Value { get; set; }
	}
}
