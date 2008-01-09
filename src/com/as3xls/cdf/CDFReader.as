package com.as3xls.cdf
{
	import flash.utils.ByteArray;
	import flash.utils.Endian;
	
	
	/**
	 * <p>
	 * Starting with Excel '97 xls documents are actually CDF (Common Document Format) files, which can also contain
	 * embedded OLE objects. Their design roughly parallels that of a file system: it starts with a header region,
	 * followed by a sector allocation table and streams (which are like files) spread out over multiple sectors.
	 * Stream names are stored in a directory stream.
	 * </p>
	 * 
	 * <p>
	 * The Workbook is stored in a stream called either Workbook or Book, depending (I believe) on the version. Also,
	 * in every Excel file I've seen this stream is conveniently pointed to by the first directory entry. Sometimes I think
	 * Microsoft does love me after all.
	 * </p>
	 * 
	 */
	public class CDFReader
	{
		private var stream:ByteArray;
		
		private var sectorSize:uint;
		private var sat:Array;
		public var dir:Array;
		
		
		/**
		 * Determines whether a ByteArray contains a CDF file by checking for the presence of the CDF magic value.
		 *  
		 * @param stream The ByteArray to check for CDF-ness
		 * @return True if the file is a CDF file; otherwise false
		 * 
		 */
		public static function isCDFFile(stream:ByteArray):Boolean {
			var magic:Array = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]; 
			if(stream.length < magic.length) {
				return false;
			}
			
			for(var n:uint = 0; n < magic.length; n++) {
				if(magic[n] != stream[n]) {
					return false;
				}
			}
			return true;
		}
		
		
		/**
		 * Wraps the given ByteArray in a CDFReader. It will be rewound and set to LittleEndian.
		 *  
		 * @param stream The ByteArray to wrap
		 * 
		 */
		public function CDFReader(stream:ByteArray) {
			this.stream = stream;
			stream.position = 0;
			stream.endian = Endian.LITTLE_ENDIAN;
			
			loadDocument();
		}
		
		/**
		 * Loads the stream represented by a given directory entry.
		 *  
		 * @param dirId The dirId pointing to the stream to load
		 * @return A ByteArray containing the specified directory
		 * 
		 */
		public function loadDirectoryEntry(dirId:uint):ByteArray {
			return loadStream(dir[dirId].secId);
		}
		
		
		/**
		 * Processes the document and arranges it so that streams can easily be extracted later 
		 * 
		 */
		private function loadDocument():void {
			// Process the header			
			var magic:Number = stream.readDouble();
			stream.position += 16; // Skip past UID
			var revision:uint = stream.readUnsignedShort();
			var version:uint = stream.readUnsignedShort();
			var endianness:uint = stream.readUnsignedShort();
			sectorSize = Math.pow(2, stream.readUnsignedShort());
			var shortSectorSize:uint = Math.pow(2, stream.readUnsignedShort());
			stream.position += 10; // Not used
			var sectorsInSAT:uint = stream.readUnsignedInt();
			var dirStreamSecID:int = stream.readInt();
			stream.position += 4; // Not used
			var minStreamSize:uint = stream.readUnsignedInt();
			var shortSecSATSecID:int = stream.readInt();
			var shortSecSATSize:uint = stream.readUnsignedInt();
			var msatSecID:int = stream.readInt();
			var msatSize:uint = stream.readUnsignedInt();
			
			// Build the sector allocation table
			sat = new Array();
			for(var i:uint = 0; i < 109; i++) {
				stream.position = 76 + i*4;
				sat[i] = stream.readInt();
			}
			
			var sectAlloc:ByteArray = new ByteArray();
			sectAlloc.endian = Endian.LITTLE_ENDIAN;
			for(i = 0; i < sat.length; i++) {
				if(sat[i] >= 0) {
					stream.position = sectorOffset(sat[i]);
					stream.readBytes(sectAlloc, i * sectorSize, sectorSize);
				} 
			}
			
			sat = new Array();
			for(i = 0; i < sectAlloc.length / 4; i++) {
				sat[i] = sectAlloc.readInt();
			}
			
			// Now load the directory
			dir = new Array();
			var directory:ByteArray = loadStream(dirStreamSecID);
			while(directory.bytesAvailable > 0) {
				var name:String = "";
				for(i = 0; i < 32; i++) {
					name += String.fromCharCode(directory.readUnsignedByte());
					directory.position++;
				}
				var nameLen:uint = directory.readUnsignedShort();
				name = name.substr(0, (nameLen-2)/2);
				var entryType:uint = directory.readUnsignedByte();
				var nodeColor:uint = directory.readUnsignedByte();
				var leftDirID:int = directory.readInt();
				var rightDirID:int = directory.readInt();
				var rootDirID:int = directory.readInt();
				directory.position += 16; // Skip past uid
				directory.position += 4; // Skip past user flags
				directory.position += 8; // Skip past timestamp
				directory.position += 8; // Skip past timestamp
				var secId:int = directory.readInt();
				var size:uint = directory.readUnsignedInt();
				directory.position += 4; // Not used
				
				if(entryType == 2) {
					dir.push({name: name, secId: secId});
				}
			}
		}
		
		/**
		 * Returns the stream at the given sector id 
		 * @param startSecID the first sector id of the stream to extract
		 * @return The stream starting at the given sector id as a ByteArray
		 * 
		 */
		private function loadStream(startSecID:uint):ByteArray {
			var ret:ByteArray = new ByteArray();
			ret.endian = Endian.LITTLE_ENDIAN;
			var secId:int = startSecID;
			var offset:uint = 0;
			
			while(secId != -2) {
				stream.position = sectorOffset(secId);
				stream.readBytes(ret, offset, sectorSize);
				offset += sectorSize;
				secId = sat[secId];
			}
			return ret;
		}

		/**
		 * Converts a sector id into an absolute offset into the raw ByteArray 
		 * @param secId The sector id to convert
		 * @return The absolute offset of the given sector id
		 * 
		 */
		private function sectorOffset(secId:uint):uint {
			return 512 + secId * sectorSize;
		}
	}
}