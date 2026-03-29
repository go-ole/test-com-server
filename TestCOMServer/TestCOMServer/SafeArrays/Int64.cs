using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000007"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayInt64
    {
        long[] EchoInt64Array1D(long[] value);
        long[,] EchoInt64Array2D(long[,] value);
        long[,,] EchoInt64Array3D(long[,,] value);
        ulong[] EchoUInt64Array1D(ulong[] value);
        ulong[,] EchoUInt64Array2D(ulong[,] value);
        ulong[,,] EchoUInt64Array3D(ulong[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000008"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Int64"),
        ComSourceInterfaces(typeof(ICOMSafeArrayInt64))]
    public class COMSafeArrayInt64 : ICOMSafeArrayInt64
    {
        public COMSafeArrayInt64() { }

        public long[] EchoInt64Array1D(long[] value) { return value; }
        public long[,] EchoInt64Array2D(long[,] value) { return value; }
        public long[,,] EchoInt64Array3D(long[,,] value) { return value; }
        public ulong[] EchoUInt64Array1D(ulong[] value) { return value; }
        public ulong[,] EchoUInt64Array2D(ulong[,] value) { return value; }
        public ulong[,,] EchoUInt64Array3D(ulong[,,] value) { return value; }
    }
}
