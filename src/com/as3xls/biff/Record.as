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
		public function readUnicodeStr16(shortLen:Boolean = false):String {
			var len:uint;
			if (shortLen) {
				len = _data.readUnsignedByte();
			} else {
				len = _data.readUnsignedShort();
			}
			var opts:uint = _data.readByte();
			var compressed:Boolean = (opts & 0x01) == 0;
			var asianPhonetic:Boolean = (opts & 0x04) == 0x04;
			var richtext:Boolean = (opts & 0x08) == 0x08;
			
			// We need to skip past these if they're present
			if (richtext) {
				_data.position += 2;
			}
			if (asianPhonetic) {
				_data.position += 4;
			}
			
			var _strArray:Array = [];
			var i:uint;
			if (compressed) {
				// This is compressed UTF-16, not UTF-8, so we don't use readUTFBytes()
				len = len > _data.bytesAvailable ? _data.bytesAvailable : len;
				for (i = 0; i < len; i++){
					_strArray.push(_data.readUnsignedByte());
				}
			} else {
				// Treating string as UCS-2, rather than UTF-16 (i.e. ignoring surrogate pairs)
				// readMultiByte() claims to do this, but doesn't seem to work...
				len = len > (_data.bytesAvailable/2) ? (_data.bytesAvailable/2) : len;
				for (i = 0; i < len; i++){
					_strArray.push(_data.readUnsignedShort());
				}
			}
			var ret:String = String.fromCharCode.apply(null, _strArray);
			
			return ret;
		}
	}
}