using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000007"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestMultiA
    {
        int GetValueA();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000008"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestMultiB
    {
        int GetValueB();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000009"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestMultiC
    {
        int GetValueC();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000A"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.Multiple"),
        ComSourceInterfaces(typeof(ICOMTestMultiA))]
    public class COMTestMultipleInterfaces : ICOMTestMultiA, ICOMTestMultiB, ICOMTestMultiC
    {
        public COMTestMultipleInterfaces() { }

        public int GetValueA() { return 10; }
        public int GetValueB() { return 20; }
        public int GetValueC() { return 30; }
    }
}
