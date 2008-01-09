package com.as3xls.biff {
	/**
	 *  Used to represent the different versions of BIFF files that exist in the wild 
	 */
	public class BIFFVersion {
		/**
		* Used by Excel 2.x. It doesn't support multiple sheets, charts, or really anything even remotely fun.
		*/
		public static const BIFF2:uint = 0;
		
		/**
		* 
		*/
		public static const BIFF3:uint = 1;
		
		
		/**
		* Used by Excel 4.0. Provides support for multiple sheets indirectly via workspaces.
		*/		
		public static const BIFF4:uint = 2;
		
		
		/**
		* Used by Excel 5 and '95. Provides support for multiple sheets natively.
		*/		
		public static const BIFF5:uint = 3;
		
		
		
		/**
		 * Used by Excel'97-2003. Generally the stream is wrapped in a CDF file.
		 * 
		 * 
		 * @see com.as3xls.cdf.CDFReader
		*/		
		public static const BIFF8:uint = 4;
	}
}