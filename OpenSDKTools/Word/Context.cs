using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools.Word
{
	class Context
	{
		public Context()
		{
			this.AbstructNumberingId = 1024;
			this.NumberingId = 1024;
			this.ImageId = 1024;
		}

		// document
		public WordprocessingDocument Package { get; set; }
		
		//
		public int AbstructNumberingId { get; set; }
		public int NumberingId { get; set; }
		public int ImageId { get; set; }

		// markers
		public List<Marker> Markers { get; set; }
		
		// template
		public Dictionary<string, Container> Containers { get; set; }
		public Dictionary<string, Variable> Variables { get; set; }


	}
}
