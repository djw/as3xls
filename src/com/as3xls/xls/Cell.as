package com.as3xls.xls {
	import flash.utils.ByteArray;
	import com.as3xls.xls.formula.Formula;
	
	public class Cell {
		private var _value:*;
		private var _note:String;
		private var _formula:Formula;
		private var _format:String;
		public var dateMode:uint = 1;
		
		public function Cell() {
			_value = "";
			_note = "";
			_format = "";
		}
		
		/**
		 * Returns the Cell's value as a native ActionScript type. If the Cell contains a formula the result of evaluating
		 * that formula will be returned.
		 * 
		 * @return A Date/Number/String depending on the Cell's type.
		 * 
		 */
		[Default]
		public function get value():* { 
			var ret:*;
			ret = _formula != null ? _formula.result : _value;
			
			if(Formatter.isDate(_format)) {
				// Handle the bizarre date system
				var d:Date = new Date();
				d.hours = 0;
				d.minutes = 0;
				d.seconds = 0;
				d.milliseconds = 0;
				d.fullYear = dateMode == ExcelFile.BASE1899 ? 1899 : 1904;
				d.month = dateMode == ExcelFile.BASE1899 ? 11: 0;
				d.date = dateMode == ExcelFile.BASE1899 ? 30 : 0;
				d.date += _value + 1;
				ret = d;
			} 
			return ret; 
		}
		public function set value(value:*):void { _value = value; }
		
		public function get note():String { return _note; }
		public function set note(value:String):void { _note = value; }
		
		public function get formula():Formula { return _formula; }
		public function set formula(value:Formula):void { _formula = value; }

		/**
		 * In shared formulas only the first formula's tokens are stored; the rest contain only a pointer to that formula.
		 * Consequently, to evaluate a shared formula the token array of the base formula needs to be available. Hence this
		 * little fella.
		*/
		public var sharedTokens:ByteArray;

		public function get format():String { return _format; }
		public function set format(value:String):void { _format = value; }

		public function toString():String {
			return value;
		}
	}
}