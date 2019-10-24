using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	internal interface IWritable
	{
		IWriter GetWriter();
	}
}
