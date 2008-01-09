package com.as3xls.xls.style {

	/**
	 * Represents and extended format record
	 */
	public class XFormat {
		private var _type:uint;
		private var _format:uint;
		
		public function XFormat(type:uint, format:uint) {
			_type = type;
			_format = format;
		}
		
		public function get type():uint { return _type; }
		public function get format():uint { return _format; }

	}
}