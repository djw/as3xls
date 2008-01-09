package com.as3xls.biff {
	import flash.utils.ByteArray;
	import flash.utils.Endian;
	
	public class BIFFWriter {
		private var _stream:ByteArray;
		
		public function BIFFWriter() {
			_stream = new ByteArray();
			_stream.endian = Endian.LITTLE_ENDIAN;
		}

		public function writeTag(r:Record):void {
			_stream.writeShort(r.type);
			_stream.writeShort(r.length);
			r.data.position = 0;
			r.data.readBytes(_stream, _stream.position, r.data.bytesAvailable);
			_stream.position += r.data.length;
		}
		
		public function get stream():ByteArray { return _stream; }
	}
}