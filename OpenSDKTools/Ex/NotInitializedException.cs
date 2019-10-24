using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class NotInitializedException : Exception
	{
		public NotInitializedException(string message)
			: base(message)
		{

		}

		public NotInitializedException(string message, Exception inner)
			: base(message, inner)
		{

		}

	}
}
