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

        string EchoString(string value);
        int PutString(string value);
        string GetString();
    }

    [ComVisible(true),
        Guid("DAA3F9FA-761E-4976-A860-8364CE55F6FC"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt16
    {
        Int16 Int16Field
        {
            get;
            set;
        }

        Int16 EchoInt16(Int16 value);
        int PutInt16(Int16 value);
        Int16 GetInt16();
    }

    [ComVisible(true),
        Guid("E3DEDEE7-38A2-4540-91D1-2EEF1D8891B0"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt32
    {
        Int32 Int32Field
        {
            get;
            set;
        }

        Int32 EchoInt32(Int32 value);
        int PutInt32(Int32 value);
        Int32 GetInt32();
    }

    [ComVisible(true),
        Guid("8D437CBC-B3ED-485C-BC32-C336432A1623"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt64
    {
        Int64 Int64Field
        {
            get;
            set;
        }

        Int64 EchoInt64(Int64 value);
        int PutInt64(Int64 value);
        Int64 GetInt64();
    }

    [ComVisible(true),
        Guid("865B85C5-0334-4AC6-9EF6-AACEC8FC5E86"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ComSourceInterfaces(
            typeof(ICOMTestString),
            typeof(ICOMTestInt16),
            typeof(ICOMTestInt32),
            typeof(ICOMTestInt64)
        )]
    public class COMTestClass1 : ICOMTestString, ICOMTestInt16, ICOMTestInt32, ICOMTestInt64
    {
        protected string rawString;
        protected Int16 rawInt16;
        protected Int32 rawInt32;
        protected Int64 rawInt64;

        public string StringField
        {
            get { return rawString; }
            set { rawString = value; }
        }

        public Int16 Int16Field
        {
            get { return rawInt16; }
            set { rawInt16 = value; }
        }

        public Int32 Int32Field
        {
            get { return rawInt32; }
            set { rawInt32 = value; }
        }

        public Int64 Int64Field
        {
            get { return rawInt64; }
            set { rawInt64 = value; }
        }

        COMTestClass1()
        {
            rawString = "";
            rawInt16 = 0;
            rawInt32 = 0;
            rawInt64 = 0;
        }

        public string EchoString(string value)
        {
            return value;
        }

        public int PutString(string value)
        {
            rawString = value;
            return 1;
        }

        public string GetString()
        {
            return rawString;
        }

        public Int16 EchoInt16(Int16 value)
        {
            return value;
        }

        public int PutInt16(Int16 value)
        {
            rawInt16 = value;
            return 1;
        }

        public Int16 GetInt16()
        {
            return rawInt16;
        }


        public Int32 EchoInt32(Int32 value)
        {
            return value;
        }

        public int PutInt32(Int32 value)
        {
            rawInt32 = value;
            return 1;
        }

        public Int32 GetInt32()
        {
            return rawInt32;
        }

        public Int64 EchoInt64(Int64 value)
        {
            return value;
        }

        public int PutInt64(Int64 value)
        {
            rawInt64 = value;
            return 1;
        }

        public Int64 GetInt64()
        {
            return rawInt64;
        }

    }
}
