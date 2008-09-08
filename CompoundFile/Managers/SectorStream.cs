using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Nix.CompoundFile.Managers
{
	public class SectorStream : ISector
	{
		private Stream baseStream;
		
		private int positionInStream = 0;
		
		private int sectorSize = 0;

		private byte defaultValue = 0;

		public SectorStream ( Stream baseStream, int positionInStream, int sectorSize, byte defaultValue )
		{
			this.baseStream = baseStream;
			this.positionInStream = positionInStream;
			this.sectorSize = sectorSize;
			this.defaultValue = defaultValue;
		}

		public void Write ( Stream stream )
		{
			baseStream.Seek(positionInStream, SeekOrigin.Begin);
			byte[] block = new byte[this.sectorSize];
			int size = baseStream.Read(block, 0, this.sectorSize);
			if ( size < this.sectorSize)
				for(int i = size; i < this.sectorSize; i++)
					block[i] = defaultValue;
			stream.Write(block, 0, this.sectorSize);
			
		}
	}
}
