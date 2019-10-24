using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	class Break : IText
	{
		internal BreakType? Type;

		internal Break(BreakType? type)
		{
			Type = type;
		}
	}
}
