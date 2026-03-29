using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000000D"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestString
    {
        string StringField { get; set; }
        int PutString(string value);
        string GetString();
        string EchoString(string value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000000E"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.String"),
        ComSourceInterfaces(typeof(ICOMTestString))]
    public class COMTestString : ICOMTestString
    {
        protected string rawString;

        public COMTestString()
        {
            rawString = "";
        }

        public string StringField
        {
            get { return rawString; }
            set { rawString = value; }
        }

        public int PutString(string value) { rawString = value; return value.Length; }
        public string GetString() { return rawString; }
        public string EchoString(string value) { return value; }
    }
}
