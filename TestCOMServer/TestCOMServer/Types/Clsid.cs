using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000021"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestClsid
    {
        string ClsidField { get; set; }
        int PutClsid(string value);
        string GetClsid();
        string EchoClsid(string value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000022"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Clsid"),
        ComSourceInterfaces(typeof(ICOMTestClsid))]
    public class COMTestClsid : ICOMTestClsid
    {
        protected Guid rawClsid;

        public COMTestClsid()
        {
            rawClsid = Guid.Empty;
        }

        public string ClsidField
        {
            get { return rawClsid.ToString("B"); }
            set { rawClsid = new Guid(value); }
        }

        public int PutClsid(string value) { rawClsid = new Guid(value); return 1; }
        public string GetClsid() { return rawClsid.ToString("B"); }
        public string EchoClsid(string value) { return value; }
    }
}
