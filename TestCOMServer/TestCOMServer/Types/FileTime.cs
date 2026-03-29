using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000025"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestFileTime
    {
        long FileTimeField { get; set; }
        int PutFileTime(long value);
        long GetFileTime();
        long EchoFileTime(long value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000026"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.FileTime"),
        ComSourceInterfaces(typeof(ICOMTestFileTime))]
    public class COMTestFileTime : ICOMTestFileTime
    {
        protected long rawFileTime;

        public COMTestFileTime()
        {
            rawFileTime = 0;
        }

        public long FileTimeField
        {
            get { return rawFileTime; }
            set { rawFileTime = value; }
        }

        public int PutFileTime(long value) { rawFileTime = value; return 1; }
        public long GetFileTime() { return rawFileTime; }
        public long EchoFileTime(long value) { return value; }
    }
}
