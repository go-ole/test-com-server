using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000003"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayInt16
    {
        short[] EchoInt16Array1D(short[] value);
        short[,] EchoInt16Array2D(short[,] value);
        short[,,] EchoInt16Array3D(short[,,] value);
        ushort[] EchoUInt16Array1D(ushort[] value);
        ushort[,] EchoUInt16Array2D(ushort[,] value);
        ushort[,,] EchoUInt16Array3D(ushort[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000004"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Int16"),
        ComSourceInterfaces(typeof(ICOMSafeArrayInt16))]
    public class COMSafeArrayInt16 : ICOMSafeArrayInt16
    {
        public COMSafeArrayInt16() { }

        public short[] EchoInt16Array1D(short[] value) { return value; }
        public short[,] EchoInt16Array2D(short[,] value) { return value; }
        public short[,,] EchoInt16Array3D(short[,,] value) { return value; }
        public ushort[] EchoUInt16Array1D(ushort[] value) { return value; }
        public ushort[,] EchoUInt16Array2D(ushort[,] value) { return value; }
        public ushort[,,] EchoUInt16Array3D(ushort[,,] value) { return value; }
    }
}
