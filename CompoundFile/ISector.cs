using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Nix.CompoundFile
{
    public interface ISector
    {
        void Write(Stream stream);
        Task WriteAsync(Stream stream, CancellationToken cancellationToken = default);
    }
}
