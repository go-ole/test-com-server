using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000003"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt16
    {
        short Int16Field { get; set; }
        ushort UInt16Field { get; set; }
        int PutInt16(short value);
        short GetInt16();
        short EchoInt16(short value);
        int PutUInt16(ushort value);
        ushort GetUInt16();
        ushort EchoUInt16(ushort value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000004"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Int16"),
        ComSourceInterfaces(typeof(ICOMTestInt16))]
    public class COMTestInt16 : ICOMTestInt16
    {
        protected short rawInt16;
        protected ushort rawUInt16;

        public COMTestInt16()
        {
            rawInt16 = 0;
            rawUInt16 = 0;
        }

        public short Int16Field
        {
            get { return rawInt16; }
            set { rawInt16 = value; }
        }

        public ushort UInt16Field
        {
            get { return rawUInt16; }
            set { rawUInt16 = value; }
        }

        public int PutInt16(short value) { rawInt16 = value; return 1; }
        public short GetInt16() { return rawInt16; }
        public short EchoInt16(short value) { return value; }

        public int PutUInt16(ushort value) { rawUInt16 = value; return 1; }
        public ushort GetUInt16() { return rawUInt16; }
        public ushort EchoUInt16(ushort value) { return value; }
    }
}
