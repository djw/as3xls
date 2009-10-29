package com.as3xls.cdf
{
	public class Directory
	{
		public var name:String;
		public var secId:int;
		public var size:uint;
		
		public function Directory(name:String, secId:int, size:uint)
		{
			this.name = name;
			this.secId = secId;
			this.size = size;
		}
	}
}