using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000017"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayVariant
    {
        object[] EchoVariantArray1D(object[] value);
        object[,] EchoVariantArray2D(object[,] value);
        object[,,] EchoVariantArray3D(object[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000018"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Variant"),
        ComSourceInterfaces(typeof(ICOMSafeArrayVariant))]
    public class COMSafeArrayVariant : ICOMSafeArrayVariant
    {
        public COMSafeArrayVariant() { }

        public object[] EchoVariantArray1D(object[] value) { return value; }
        public object[,] EchoVariantArray2D(object[,] value) { return value; }
        public object[,,] EchoVariantArray3D(object[,,] value) { return value; }
    }
}
