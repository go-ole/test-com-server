using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000005"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt32
    {
        int Int32Field { get; set; }
        uint UInt32Field { get; set; }
        int PutInt32(int value);
        int GetInt32();
        int EchoInt32(int value);
        int PutUInt32(uint value);
        uint GetUInt32();
        uint EchoUInt32(uint value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000006"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Int32"),
        ComSourceInterfaces(typeof(ICOMTestInt32))]
    public class COMTestInt32 : ICOMTestInt32
    {
        protected int rawInt32;
        protected uint rawUInt32;

        public COMTestInt32()
        {
            rawInt32 = 0;
            rawUInt32 = 0;
        }

        public int Int32Field
        {
            get { return rawInt32; }
            set { rawInt32 = value; }
        }

        public uint UInt32Field
        {
            get { return rawUInt32; }
            set { rawUInt32 = value; }
        }

        public int PutInt32(int value) { rawInt32 = value; return 1; }
        public int GetInt32() { return rawInt32; }
        public int EchoInt32(int value) { return value; }

        public int PutUInt32(uint value) { rawUInt32 = value; return 1; }
        public uint GetUInt32() { return rawUInt32; }
        public uint EchoUInt32(uint value) { return value; }
    }
}
