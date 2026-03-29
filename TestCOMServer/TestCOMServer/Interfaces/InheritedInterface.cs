using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000B"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestBaseInterface
    {
        int GetBaseValue();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000C"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestDerivedInterface : ICOMTestBaseInterface
    {
        int GetDerivedValue();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000D"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.Inherited"),
        ComSourceInterfaces(typeof(ICOMTestDerivedInterface))]
    public class COMTestInheritedInterface : ICOMTestDerivedInterface
    {
        public COMTestInheritedInterface() { }

        public int GetBaseValue() { return 100; }
        public int GetDerivedValue() { return 200; }
    }
}
