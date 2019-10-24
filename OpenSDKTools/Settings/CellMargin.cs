using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools.Settings
{
	public class CellMargin
	{
		public bool TopMarginSet { get; private set; } = false;
		public bool BottomMarginSet { get; private set; } = false;
		public bool LeftMarginSet { get; private set; } = false;
		public bool RightMarginSet { get; private set; } = false;

		public bool MarginSet
		{
			get
			{
				return TopMarginSet || RightMarginSet || BottomMarginSet || LeftMarginSet;
			}
		}

		private int _top;
		public int Top
		{ get
			{
				return _top;
			}

			set
			{
				_top = value;
				TopMarginSet = true;
			}
		}

		private int _bottom;
		public int Bottom
		{
			get
			{
				return _bottom;
			}
			set
			{
				_bottom = value;
				BottomMarginSet = true;
			}
		}

		private int _left;
		public int Left
		{
			get
			{
				return _left;
			}
			set
			{
				_left = value;
				LeftMarginSet = true;
			}
		}

		private int _right;
		public int Right
		{
			get
			{
				return _right;
			}
			set
			{
				_right = value;
				RightMarginSet = true;
			}
		}


	}
}
