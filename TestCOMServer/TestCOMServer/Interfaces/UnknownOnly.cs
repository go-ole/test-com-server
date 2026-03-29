using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000003"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ICOMTestUnknownOnly
    {
        int GetValue();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000004"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.UnknownOnly"),
        ComSourceInterfaces(typeof(ICOMTestUnknownOnly))]
    public class COMTestUnknownOnly : ICOMTestUnknownOnly
    {
        public COMTestUnknownOnly() { }

        public int GetValue() { return 84; }
    }
}
