using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Nix.CompoundFile
{
	public interface ISector
	{
		void Write(Stream stream);
	}
}
