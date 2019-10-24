using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools.Word
{
	public class Variable
	{
		public string Key { get; set; }
		public string Value { get; set; }

		internal IWriter GetWriter()
		{
			return new Word.VariableWriter();
		}
	}
}
