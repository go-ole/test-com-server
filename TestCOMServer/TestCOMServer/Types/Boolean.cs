using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000000F"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestBoolean
    {
        bool BooleanField { get; set; }
        int PutBoolean(bool value);
        bool GetBoolean();
        bool EchoBoolean(bool value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000010"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Boolean"),
        ComSourceInterfaces(typeof(ICOMTestBoolean))]
    public class COMTestBoolean : ICOMTestBoolean
    {
        protected bool rawBoolean;

        public COMTestBoolean()
        {
            rawBoolean = false;
        }

        public bool BooleanField
        {
            get { return rawBoolean; }
            set { rawBoolean = value; }
        }

        public int PutBoolean(bool value) { rawBoolean = value; return 1; }
        public bool GetBoolean() { return rawBoolean; }
        public bool EchoBoolean(bool value) { return value; }
    }
}
