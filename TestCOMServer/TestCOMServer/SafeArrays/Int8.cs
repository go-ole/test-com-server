using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000001"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayInt8
    {
        sbyte[] EchoInt8Array1D(sbyte[] value);
        sbyte[,] EchoInt8Array2D(sbyte[,] value);
        sbyte[,,] EchoInt8Array3D(sbyte[,,] value);
        byte[] EchoUInt8Array1D(byte[] value);
        byte[,] EchoUInt8Array2D(byte[,] value);
        byte[,,] EchoUInt8Array3D(byte[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000002"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Int8"),
        ComSourceInterfaces(typeof(ICOMSafeArrayInt8))]
    public class COMSafeArrayInt8 : ICOMSafeArrayInt8
    {
        public COMSafeArrayInt8() { }

        public sbyte[] EchoInt8Array1D(sbyte[] value) { return value; }
        public sbyte[,] EchoInt8Array2D(sbyte[,] value) { return value; }
        public sbyte[,,] EchoInt8Array3D(sbyte[,,] value) { return value; }
        public byte[] EchoUInt8Array1D(byte[] value) { return value; }
        public byte[,] EchoUInt8Array2D(byte[,] value) { return value; }
        public byte[,,] EchoUInt8Array3D(byte[,,] value) { return value; }
    }
}
