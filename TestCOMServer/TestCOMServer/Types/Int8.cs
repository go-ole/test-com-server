using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000001"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt8
    {
        sbyte Int8Field { get; set; }
        byte UInt8Field { get; set; }
        int PutInt8(sbyte value);
        sbyte GetInt8();
        sbyte EchoInt8(sbyte value);
        int PutUInt8(byte value);
        byte GetUInt8();
        byte EchoUInt8(byte value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000002"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Int8"),
        ComSourceInterfaces(typeof(ICOMTestInt8))]
    public class COMTestInt8 : ICOMTestInt8
    {
        protected sbyte rawInt8;
        protected byte rawUInt8;

        public COMTestInt8()
        {
            rawInt8 = 0;
            rawUInt8 = 0;
        }

        public sbyte Int8Field
        {
            get { return rawInt8; }
            set { rawInt8 = value; }
        }

        public byte UInt8Field
        {
            get { return rawUInt8; }
            set { rawUInt8 = value; }
        }

        public int PutInt8(sbyte value) { rawInt8 = value; return 1; }
        public sbyte GetInt8() { return rawInt8; }
        public sbyte EchoInt8(sbyte value) { return value; }

        public int PutUInt8(byte value) { rawUInt8 = value; return 1; }
        public byte GetUInt8() { return rawUInt8; }
        public byte EchoUInt8(byte value) { return value; }
    }
}
