using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000F"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayBoolean
    {
        bool[] EchoBooleanArray1D(bool[] value);
        bool[,] EchoBooleanArray2D(bool[,] value);
        bool[,,] EchoBooleanArray3D(bool[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000010"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Boolean"),
        ComSourceInterfaces(typeof(ICOMSafeArrayBoolean))]
    public class COMSafeArrayBoolean : ICOMSafeArrayBoolean
    {
        public COMSafeArrayBoolean() { }

        public bool[] EchoBooleanArray1D(bool[] value) { return value; }
        public bool[,] EchoBooleanArray2D(bool[,] value) { return value; }
        public bool[,,] EchoBooleanArray3D(bool[,,] value) { return value; }
    }
}
