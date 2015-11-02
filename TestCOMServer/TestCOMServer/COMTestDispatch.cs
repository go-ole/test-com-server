using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("E0133EB4-C36F-469A-9D3D-C66B84BE19ED"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestString
    {
        string StringField
        {
            get;
            set;
        }
        
        int PutString(string value);
        string GetString();
    }

    [ComVisible(true),
        Guid("BEB06610-EB84-4155-AF58-E2BFF53608B4"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt8
    {
        sbyte Int8Field
        {
            get;
            set;
        }

        byte UInt8Field
        {
            get;
            set;
        }

        int PutInt8(sbyte value);
        sbyte GetInt8();

        int PutUInt8(byte value);
        byte GetUInt8();
    }

    [ComVisible(true),
        Guid("DAA3F9FA-761E-4976-A860-8364CE55F6FC"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt16
    {
        short Int16Field
        {
            get;
            set;
        }

        ushort UInt16Field
        {
            get;
            set;
        }

        int PutInt16(short value);
        short GetInt16();

        int PutUInt16(ushort value);
        ushort GetUInt16();
    }

    [ComVisible(true),
        Guid("E3DEDEE7-38A2-4540-91D1-2EEF1D8891B0"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt32
    {
        int Int32Field
        {
            get;
            set;
        }

        uint UInt32Field
        {
            get;
            set;
        }

        int PutInt32(int value);
        int GetInt32();

        int PutUInt32(uint value);
        uint GetUInt32();
    }

    [ComVisible(true),
        Guid("8D437CBC-B3ED-485C-BC32-C336432A1623"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt64
    {
        long Int64Field
        {
            get;
            set;
        }

        ulong UInt64Field
        {
            get;
            set;
        }
        
        int PutInt64(long value);
        long GetInt64();
        
        int PutUInt64(ulong value);
        ulong GetUInt64();
    }

    [ComVisible(true),
        Guid("BF1ED004-EA02-456A-AA55-2AC8AC6B054C"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestFloat
    {
        float Float32Field
        {
            get;
            set;
        }
        
        int PutFloat32(float value);
        float GetFloat32();
    }

    [ComVisible(true),
        Guid("BF908A81-8687-4E93-999F-D86FAB284BA0"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestDouble
    {
        double Float64Field
        {
            get;
            set;
        }
        
        int PutFloat64(double value);
        double GetFloat64();
    }

    [ComVisible(true),
        Guid("D530E7A6-4EE8-40D1-8931-3D63B8605001"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestBoolean
    {
        bool BooleanField
        {
            get;
            set;
        }
        
        int PutBoolean(bool value);
        bool GetBoolean();
    }

    [ComVisible(true),
        Guid("CCA8D7AE-91C0-4277-A8B3-FF4EDF28D3C0"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestTypes : 
        ICOMTestString,
        ICOMTestInt8,
        ICOMTestInt16,
        ICOMTestInt32,
        ICOMTestInt64,
        ICOMTestFloat,
        ICOMTestDouble,
        ICOMTestBoolean
    { }

    [ComVisible(true),
        Guid("6485B1EF-D780-4834-A4FE-1EBB51746CA3"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMEchoTestObject
    {
        sbyte EchoInt8(sbyte value);
        byte EchoUInt8(byte value);
        short EchoInt16(short value);
        ushort EchoUInt16(ushort value);
        int EchoInt32(int value);
        uint EchoUInt32(uint value);
        long EchoInt64(long value);
        ulong EchoUInt64(ulong value);
        float EchoFloat32(float value);
        double EchoFloat64(double value);
        string EchoString(string value);
        bool EchoBoolean(bool value);
    }
    
    [ComVisible(true),
        Guid("3C24506A-AE9E-4D50-9157-EF317281F1B0"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ComSourceInterfaces(typeof(ICOMEchoTestObject))]
    public class COMEchoTestObject : ICOMEchoTestObject
    {
        public COMEchoTestObject()
        {
        }

        public String EchoString(String value)
        {
            return value;
        }

        public sbyte EchoInt8(sbyte value)
        {
            return value;
        }

        public byte EchoUInt8(byte value)
        {
            return value;
        }

        public short EchoInt16(short value)
        {
            return value;
        }

        public ushort EchoUInt16(ushort value)
        {
            return value;
        }

        public int EchoInt32(int value)
        {
            return value;
        }

        public uint EchoUInt32(uint value)
        {
            return value;
        }

        public long EchoInt64(long value)
        {
            return value;
        }

        public ulong EchoUInt64(ulong value)
        {
            return value;
        }

        public float EchoFloat32(float value)
        {
            return value;
        }

        public double EchoFloat64(double value)
        {
            return value;
        }

        public bool EchoBoolean(bool value)
        {
            return value;
        }
    }

    [ComVisible(true),
        Guid("865B85C5-0334-4AC6-9EF6-AACEC8FC5E86"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ComSourceInterfaces(typeof(ICOMTestTypes))]
    public class COMTestScalarClass : ICOMTestTypes
    {
        protected string rawString;
        protected sbyte rawInt8;
        protected byte rawUInt8;
        protected short rawInt16;
        protected ushort rawUInt16;
        protected int rawInt32;
        protected uint rawUInt32;
        protected long rawInt64;
        protected ulong rawUInt64;
        protected float rawFloat;
        protected double rawDouble;
        protected bool rawBoolean;

        public string StringField
        {
            get { return rawString; }
            set { rawString = value; }
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

        public long Int64Field
        {
            get { return rawInt64; }
            set { rawInt64 = value; }
        }

        public ulong UInt64Field
        {
            get { return rawUInt64; }
            set { rawUInt64 = value; }
        }

        public float Float32Field
        {
            get { return rawFloat; }
            set { rawFloat = value; }
        }

        public double Float64Field
        {
            get { return rawDouble; }
            set { rawDouble = value; }
        }

        public bool BooleanField
        {
            get { return rawBoolean; }
            set { rawBoolean = value; }
        }

        public COMTestScalarClass()
        {
            rawString = "";
            rawInt8 = 0;
            rawUInt8 = 0;
            rawInt16 = 0;
            rawUInt16 = 0;
            rawInt32 = 0;
            rawUInt32 = 0;
            rawInt64 = 0;
            rawUInt64 = 0;
            rawFloat = 0.0f;
            rawDouble = 0.0;
            rawBoolean = false;
        }

        public int PutString(string value)
        {
            rawString = value;
            return value.Length;
        }

        public string GetString()
        {
            return rawString;
        }

        public int PutInt8(sbyte value)
        {
            rawInt8 = value;
            return 1;
        }

        public sbyte GetInt8()
        {
            return rawInt8;
        }

        public int PutUInt8(byte value)
        {
            rawUInt8 = value;
            return 1;
        }

        public byte GetUInt8()
        {
            return rawUInt8;
        }

        public int PutInt16(short value)
        {
            rawInt16 = value;
            return 1;
        }

        public short GetInt16()
        {
            return rawInt16;
        }

        public int PutUInt16(ushort value)
        {
            rawUInt16 = value;
            return 1;
        }

        public ushort GetUInt16()
        {
            return rawUInt16;
        }

        public int PutInt32(int value)
        {
            rawInt32 = value;
            return 1;
        }

        public int GetInt32()
        {
            return rawInt32;
        }

        public int PutUInt32(uint value)
        {
            rawUInt32 = value;
            return 1;
        }

        public uint GetUInt32()
        {
            return rawUInt32;
        }

        public int PutInt64(long value)
        {
            rawInt64 = value;
            return 1;
        }

        public long GetInt64()
        {
            return rawInt64;
        }

        public int PutUInt64(ulong value)
        {
            rawUInt64 = value;
            return 1;
        }

        public ulong GetUInt64()
        {
            return rawUInt64;
        }

        public int PutFloat32(float value)
        {
            rawFloat = value;
            return 1;
        }

        public float GetFloat32()
        {
            return rawFloat;
        }

        public int PutFloat64(double value)
        {
            rawDouble = value;
            return 1;
        }

        public double GetFloat64()
        {
            return rawDouble;
        }

        public int PutBoolean(bool value)
        {
            rawBoolean = value;
            return 1;
        }

        public bool GetBoolean()
        {
            return rawBoolean;
        }

    }
}
