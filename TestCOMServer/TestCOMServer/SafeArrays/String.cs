using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000D"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayString
    {
        string[] EchoStringArray1D(string[] value);
        string[,] EchoStringArray2D(string[,] value);
        string[,,] EchoStringArray3D(string[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000E"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.String"),
        ComSourceInterfaces(typeof(ICOMSafeArrayString))]
    public class COMSafeArrayString : ICOMSafeArrayString
    {
        public COMSafeArrayString() { }

        public string[] EchoStringArray1D(string[] value) { return value; }
        public string[,] EchoStringArray2D(string[,] value) { return value; }
        public string[,,] EchoStringArray3D(string[,,] value) { return value; }
    }
}
