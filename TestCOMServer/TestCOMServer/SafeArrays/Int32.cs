using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000005"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayInt32
    {
        int[] EchoInt32Array1D(int[] value);
        int[,] EchoInt32Array2D(int[,] value);
        int[,,] EchoInt32Array3D(int[,,] value);
        uint[] EchoUInt32Array1D(uint[] value);
        uint[,] EchoUInt32Array2D(uint[,] value);
        uint[,,] EchoUInt32Array3D(uint[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000006"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Int32"),
        ComSourceInterfaces(typeof(ICOMSafeArrayInt32))]
    public class COMSafeArrayInt32 : ICOMSafeArrayInt32
    {
        public COMSafeArrayInt32() { }

        public int[] EchoInt32Array1D(int[] value) { return value; }
        public int[,] EchoInt32Array2D(int[,] value) { return value; }
        public int[,,] EchoInt32Array3D(int[,,] value) { return value; }
        public uint[] EchoUInt32Array1D(uint[] value) { return value; }
        public uint[,] EchoUInt32Array2D(uint[,] value) { return value; }
        public uint[,,] EchoUInt32Array3D(uint[,,] value) { return value; }
    }
}
