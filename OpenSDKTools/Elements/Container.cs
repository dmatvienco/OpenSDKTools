using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class Container
	{
		public Container() : this(string.Empty)
		{

		}

		public Container(string key)
		{
			this.Key = key;
			this.IsNumbering = false;
			this.Children = new List<IWritable>();
			this.NumberingId = 0;
		}

		public bool IsNumbering { get; set; }

		public string Key { get; set; }

		internal List<IWritable> Children { get; set; }
		
		internal int NumberingId { get; set; }

		public void AddChild(Paragraph paragraph)
		{
			this.Children.Add(paragraph);
		}

		public void AddChild(Table table)
		{
			this.Children.Add(table);
		}

		public void AddChild(Image image)
		{
			this.Children.Add(image);
		}

		public void AddChild(Chart chart)
		{
			this.Children.Add(chart);
		}
	}
}
