using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class NotSetException : Exception
	{
		public NotSetException(string message)
			: base(message)
		{

		}

		public NotSetException(string message, Exception inner)
			: base(message, inner)
		{

		}

	}
}
