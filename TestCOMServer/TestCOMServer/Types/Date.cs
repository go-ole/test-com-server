using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000013"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestDate
    {
        DateTime DateField { get; set; }
        int PutDate(DateTime value);
        DateTime GetDate();
        DateTime EchoDate(DateTime value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000014"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Date"),
        ComSourceInterfaces(typeof(ICOMTestDate))]
    public class COMTestDate : ICOMTestDate
    {
        protected DateTime rawDate;

        public COMTestDate()
        {
            rawDate = DateTime.MinValue;
        }

        public DateTime DateField
        {
            get { return rawDate; }
            set { rawDate = value; }
        }

        public int PutDate(DateTime value) { rawDate = value; return 1; }
        public DateTime GetDate() { return rawDate; }
        public DateTime EchoDate(DateTime value) { return value; }
    }
}
