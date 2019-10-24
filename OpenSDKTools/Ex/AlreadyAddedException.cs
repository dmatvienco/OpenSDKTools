using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class AlreadyAddedException : Exception
	{
		public AlreadyAddedException(string message)
			: base(message)
		{

		}

		public AlreadyAddedException(string message, Exception inner)
			: base(message, inner)
		{

		}

	}
}
