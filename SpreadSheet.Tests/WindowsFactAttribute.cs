using System.Runtime.InteropServices;
using Xunit;

namespace SpreadSheet.Tests
{
    public sealed class WindowsFactAttribute : FactAttribute
    {
        public WindowsFactAttribute()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                Skip = "Row height measurement uses System.Drawing.Common and requires Windows.";
        }
    }
}
