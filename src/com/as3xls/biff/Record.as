package com.as3xls.biff {
	import flash.utils.ByteArray;
	import flash.utils.Endian;
	
	
	/**
	 * Represents a single BIFF record.
	 */
	public class Record
	{
		private var _type:uint;
		private var _data:ByteArray;
		
		
		/**
		 * 
		 * @param type The type of the BIFF record
		 * @param data A byte array containing the BIFF record without the length and type header
		 * 
		 */
		public function Record(type:uint, data:ByteArray = null) {
			_type = type;
			_data = data == null ? new ByteArray() : data;
			_data.endian = Endian.LITTLE_ENDIAN;
		}
		
		public function get type():uint { return _type; }
		public function set type(value:uint):void { _type = value; }
		
		public function get data():ByteArray { return _data; }
		public function get length():uint { return _data.length; }

		/**
		 * 	Reads one of the weird Excel 2 byte length unicode strings.
		 * 
		 *  @return A unicode string with 2 byte length read from the _data ByteArray's current position.
		 */
		public function readUnicodeStr16():String {
			var len:uint = _data.readUnsignedShort();
			var opts:uint = _data.readByte();
			var compressed:Boolean = (opts & 0x01) == 0;
			var asianPhonetic:Boolean = (opts & 0x04) == 0x04;
			var richtext:Boolean = (opts & 0x08) == 0x08;
			
			if(!compressed) {
				len *= 2;
			}
			
			len = len > _data.bytesAvailable ? _data.bytesAvailable : len;
			var ret:String = _data.readUTFBytes(len);
			return ret;
		}
	}
}