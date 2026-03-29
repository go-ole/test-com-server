using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000007"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt64
    {
        long Int64Field { get; set; }
        ulong UInt64Field { get; set; }
        int PutInt64(long value);
        long GetInt64();
        long EchoInt64(long value);
        int PutUInt64(ulong value);
        ulong GetUInt64();
        ulong EchoUInt64(ulong value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000008"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Int64"),
        ComSourceInterfaces(typeof(ICOMTestInt64))]
    public class COMTestInt64 : ICOMTestInt64
    {
        protected long rawInt64;
        protected ulong rawUInt64;

        public COMTestInt64()
        {
            rawInt64 = 0;
            rawUInt64 = 0;
        }

        public long Int64Field
        {
            get { return rawInt64; }
            set { rawInt64 = value; }
        }

        public ulong UInt64Field
        {
            get { return rawUInt64; }
            set { rawUInt64 = value; }
        }

        public int PutInt64(long value) { rawInt64 = value; return 1; }
        public long GetInt64() { return rawInt64; }
        public long EchoInt64(long value) { return value; }

        public int PutUInt64(ulong value) { rawUInt64 = value; return 1; }
        public ulong GetUInt64() { return rawUInt64; }
        public ulong EchoUInt64(ulong value) { return value; }
    }
}
