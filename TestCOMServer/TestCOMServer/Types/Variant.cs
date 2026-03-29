using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000019"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestVariant
    {
        object VariantField { get; set; }
        int PutVariant(object value);
        object GetVariant();
        object EchoVariant(object value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001A"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Variant"),
        ComSourceInterfaces(typeof(ICOMTestVariant))]
    public class COMTestVariant : ICOMTestVariant
    {
        protected object rawVariant;

        public COMTestVariant()
        {
            rawVariant = null;
        }

        public object VariantField
        {
            get { return rawVariant; }
            set { rawVariant = value; }
        }

        public int PutVariant(object value) { rawVariant = value; return 1; }
        public object GetVariant() { return rawVariant; }
        public object EchoVariant(object value) { return value; }
    }
}
