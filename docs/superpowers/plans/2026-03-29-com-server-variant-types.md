# COM Server Full VARIANT Type & Interface Support Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rebuild the .NET COM test server to support all VARIANT types, multi-dimensional SAFEARRAYs, and all COM interface patterns for go-ole testing.

**Architecture:** One COM class per type with get/set/echo pattern. SafeArrays echo 1D/2D/3D arrays. Interface pattern classes return hardcoded constants. All organized under Types/, SafeArrays/, Interfaces/ directories.

**Tech Stack:** .NET 5.0 (net5.0-windows), C#, COM Interop, SDK-style project

**Note:** This is a COM server library tested externally by go-ole (Go). There are no unit tests in this project. Verification is "does it compile successfully."

---

## File Structure

**Delete:**
- `TestCOMServer/TestCOMServer/COMTestDispatch.cs`
- `TestCOMServer/TestCOMServer/COMSafeArray.cs`
- `TestCOMServer/TestCOMServer/Properties/AssemblyInfo.cs`

**Create:**
- `TestCOMServer/TestCOMServer/Types/Int8.cs`
- `TestCOMServer/TestCOMServer/Types/Int16.cs`
- `TestCOMServer/TestCOMServer/Types/Int32.cs`
- `TestCOMServer/TestCOMServer/Types/Int64.cs`
- `TestCOMServer/TestCOMServer/Types/Float32.cs`
- `TestCOMServer/TestCOMServer/Types/Float64.cs`
- `TestCOMServer/TestCOMServer/Types/String.cs`
- `TestCOMServer/TestCOMServer/Types/Boolean.cs`
- `TestCOMServer/TestCOMServer/Types/Currency.cs`
- `TestCOMServer/TestCOMServer/Types/Date.cs`
- `TestCOMServer/TestCOMServer/Types/Decimal.cs`
- `TestCOMServer/TestCOMServer/Types/Error.cs`
- `TestCOMServer/TestCOMServer/Types/Variant.cs`
- `TestCOMServer/TestCOMServer/Types/Unknown.cs`
- `TestCOMServer/TestCOMServer/Types/Dispatch.cs`
- `TestCOMServer/TestCOMServer/Types/Empty.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Int8.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Int16.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Int32.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Int64.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Float32.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Float64.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/String.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Boolean.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Currency.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Date.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Decimal.cs`
- `TestCOMServer/TestCOMServer/SafeArrays/Variant.cs`
- `TestCOMServer/TestCOMServer/Interfaces/DualInterface.cs`
- `TestCOMServer/TestCOMServer/Interfaces/UnknownOnly.cs`
- `TestCOMServer/TestCOMServer/Interfaces/DispatchOnly.cs`
- `TestCOMServer/TestCOMServer/Interfaces/MultipleInterfaces.cs`
- `TestCOMServer/TestCOMServer/Interfaces/InheritedInterface.cs`
- `TestCOMServer/TestCOMServer/Interfaces/ConnectionPoint.cs`

**Modify:**
- `TestCOMServer/TestCOMServer/TestCOMServer.csproj` (replace with SDK-style)
- `TestCOMServer/TestCOMServer.sln` (update for new project format)

---

### Task 1: Migrate Project to .NET 5.0 SDK-Style

**Files:**
- Modify: `TestCOMServer/TestCOMServer/TestCOMServer.csproj`
- Delete: `TestCOMServer/TestCOMServer/Properties/AssemblyInfo.cs`
- Delete: `TestCOMServer/TestCOMServer/COMTestDispatch.cs`
- Delete: `TestCOMServer/TestCOMServer/COMSafeArray.cs`

- [ ] **Step 1: Delete old source files**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server
rm TestCOMServer/TestCOMServer/COMTestDispatch.cs
rm TestCOMServer/TestCOMServer/COMSafeArray.cs
rm TestCOMServer/TestCOMServer/Properties/AssemblyInfo.cs
rmdir TestCOMServer/TestCOMServer/Properties
```

- [ ] **Step 2: Replace .csproj with SDK-style project**

Replace the entire contents of `TestCOMServer/TestCOMServer/TestCOMServer.csproj` with:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net5.0-windows</TargetFramework>
    <OutputType>Library</OutputType>
    <RootNamespace>TestCOMServer</RootNamespace>
    <AssemblyName>TestCOMServer</AssemblyName>
    <EnableComHosting>true</EnableComHosting>
    <GenerateAssemblyInfo>true</GenerateAssemblyInfo>
    <AssemblyVersion>2.0.0.0</AssemblyVersion>
    <FileVersion>2.0.0.0</FileVersion>
    <Version>2.0.0</Version>
    <Company>Go-OLE</Company>
    <Product>TestCOMServer</Product>
    <Copyright>Copyright Go-OLE 2015</Copyright>
  </PropertyGroup>

  <ItemGroup>
    <Using Include="System.Runtime.InteropServices" />
  </ItemGroup>

</Project>
```

- [ ] **Step 3: Create empty directory structure**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server
mkdir -p TestCOMServer/TestCOMServer/Types
mkdir -p TestCOMServer/TestCOMServer/SafeArrays
mkdir -p TestCOMServer/TestCOMServer/Interfaces
```

- [ ] **Step 4: Commit**

```bash
git add -A TestCOMServer/
git commit -m "Migrate to .NET 5.0 SDK-style project, delete old source files"
```

---

### Task 2: Scalar Types - Integer Types (Int8, Int16, Int32, Int64)

**Files:**
- Create: `TestCOMServer/TestCOMServer/Types/Int8.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Int16.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Int32.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Int64.cs`

- [ ] **Step 1: Create Types/Int8.cs**

```csharp
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
```

- [ ] **Step 2: Create Types/Int16.cs**

```csharp
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
```

- [ ] **Step 3: Create Types/Int32.cs**

```csharp
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
```

- [ ] **Step 4: Create Types/Int64.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000007"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestInt64
    {
        long Int64Field { get; set; }
        ulong UInt64Field { get; set; }
        int PutInt64(long value);
        long GetInt64();
        long EchoInt64(long value);
        int PutUInt64(ulong value);
        ulong GetUInt64();
        ulong EchoUInt64(ulong value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000008"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Int64"),
        ComSourceInterfaces(typeof(ICOMTestInt64))]
    public class COMTestInt64 : ICOMTestInt64
    {
        protected long rawInt64;
        protected ulong rawUInt64;

        public COMTestInt64()
        {
            rawInt64 = 0;
            rawUInt64 = 0;
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

        public int PutInt64(long value) { rawInt64 = value; return 1; }
        public long GetInt64() { return rawInt64; }
        public long EchoInt64(long value) { return value; }

        public int PutUInt64(ulong value) { rawUInt64 = value; return 1; }
        public ulong GetUInt64() { return rawUInt64; }
        public ulong EchoUInt64(ulong value) { return value; }
    }
}
```

- [ ] **Step 5: Verify build compiles**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 6: Commit**

```bash
git add TestCOMServer/TestCOMServer/Types/Int8.cs TestCOMServer/TestCOMServer/Types/Int16.cs TestCOMServer/TestCOMServer/Types/Int32.cs TestCOMServer/TestCOMServer/Types/Int64.cs
git commit -m "Add scalar integer types (Int8, Int16, Int32, Int64) with signed/unsigned"
```

---

### Task 3: Scalar Types - Floating Point and String (Float32, Float64, String)

**Files:**
- Create: `TestCOMServer/TestCOMServer/Types/Float32.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Float64.cs`
- Create: `TestCOMServer/TestCOMServer/Types/String.cs`

- [ ] **Step 1: Create Types/Float32.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000009"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestFloat32
    {
        float Float32Field { get; set; }
        int PutFloat32(float value);
        float GetFloat32();
        float EchoFloat32(float value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000000A"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Float32"),
        ComSourceInterfaces(typeof(ICOMTestFloat32))]
    public class COMTestFloat32 : ICOMTestFloat32
    {
        protected float rawFloat32;

        public COMTestFloat32()
        {
            rawFloat32 = 0.0f;
        }

        public float Float32Field
        {
            get { return rawFloat32; }
            set { rawFloat32 = value; }
        }

        public int PutFloat32(float value) { rawFloat32 = value; return 1; }
        public float GetFloat32() { return rawFloat32; }
        public float EchoFloat32(float value) { return value; }
    }
}
```

- [ ] **Step 2: Create Types/Float64.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000000B"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestFloat64
    {
        double Float64Field { get; set; }
        int PutFloat64(double value);
        double GetFloat64();
        double EchoFloat64(double value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000000C"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Float64"),
        ComSourceInterfaces(typeof(ICOMTestFloat64))]
    public class COMTestFloat64 : ICOMTestFloat64
    {
        protected double rawFloat64;

        public COMTestFloat64()
        {
            rawFloat64 = 0.0;
        }

        public double Float64Field
        {
            get { return rawFloat64; }
            set { rawFloat64 = value; }
        }

        public int PutFloat64(double value) { rawFloat64 = value; return 1; }
        public double GetFloat64() { return rawFloat64; }
        public double EchoFloat64(double value) { return value; }
    }
}
```

- [ ] **Step 3: Create Types/String.cs**

```csharp
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
```

- [ ] **Step 4: Verify build compiles**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 5: Commit**

```bash
git add TestCOMServer/TestCOMServer/Types/Float32.cs TestCOMServer/TestCOMServer/Types/Float64.cs TestCOMServer/TestCOMServer/Types/String.cs
git commit -m "Add scalar types: Float32, Float64, String"
```

---

### Task 4: Scalar Types - Boolean, Currency, Date, Decimal

**Files:**
- Create: `TestCOMServer/TestCOMServer/Types/Boolean.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Currency.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Date.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Decimal.cs`

- [ ] **Step 1: Create Types/Boolean.cs**

```csharp
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
```

- [ ] **Step 2: Create Types/Currency.cs**

VT_CY is a 64-bit integer scaled by 10,000. In .NET COM interop, `decimal` with `MarshalAs(UnmanagedType.Currency)` marshals as VT_CY.

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000011"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestCurrency
    {
        [return: MarshalAs(UnmanagedType.Currency)]
        decimal CurrencyField
        {
            [return: MarshalAs(UnmanagedType.Currency)]
            get;
            [param: MarshalAs(UnmanagedType.Currency)]
            set;
        }

        int PutCurrency([MarshalAs(UnmanagedType.Currency)] decimal value);

        [return: MarshalAs(UnmanagedType.Currency)]
        decimal GetCurrency();

        [return: MarshalAs(UnmanagedType.Currency)]
        decimal EchoCurrency([MarshalAs(UnmanagedType.Currency)] decimal value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000012"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Currency"),
        ComSourceInterfaces(typeof(ICOMTestCurrency))]
    public class COMTestCurrency : ICOMTestCurrency
    {
        protected decimal rawCurrency;

        public COMTestCurrency()
        {
            rawCurrency = 0m;
        }

        public decimal CurrencyField
        {
            get { return rawCurrency; }
            set { rawCurrency = value; }
        }

        public int PutCurrency(decimal value) { rawCurrency = value; return 1; }
        public decimal GetCurrency() { return rawCurrency; }
        public decimal EchoCurrency(decimal value) { return value; }
    }
}
```

- [ ] **Step 3: Create Types/Date.cs**

VT_DATE is an OLE Automation date (double). .NET DateTime marshals as VT_DATE in COM interop automatically.

```csharp
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
```

- [ ] **Step 4: Create Types/Decimal.cs**

VT_DECIMAL is a 96-bit decimal. .NET `decimal` marshals as VT_DECIMAL when used as a VARIANT. Note: VT_DECIMAL cannot be a top-level VARIANT type on its own in all COM contexts, but it works inside VARIANT (VT_BYREF|VT_DECIMAL). We use `object` to force VARIANT wrapping.

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000015"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestDecimal
    {
        object DecimalField { get; set; }
        int PutDecimal(object value);
        object GetDecimal();
        object EchoDecimal(object value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000016"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Decimal"),
        ComSourceInterfaces(typeof(ICOMTestDecimal))]
    public class COMTestDecimal : ICOMTestDecimal
    {
        protected decimal rawDecimal;

        public COMTestDecimal()
        {
            rawDecimal = 0m;
        }

        public object DecimalField
        {
            get { return rawDecimal; }
            set { rawDecimal = (decimal)value; }
        }

        public int PutDecimal(object value) { rawDecimal = (decimal)value; return 1; }
        public object GetDecimal() { return rawDecimal; }
        public object EchoDecimal(object value) { return value; }
    }
}
```

- [ ] **Step 5: Verify build compiles**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 6: Commit**

```bash
git add TestCOMServer/TestCOMServer/Types/Boolean.cs TestCOMServer/TestCOMServer/Types/Currency.cs TestCOMServer/TestCOMServer/Types/Date.cs TestCOMServer/TestCOMServer/Types/Decimal.cs
git commit -m "Add scalar types: Boolean, Currency, Date, Decimal"
```

---

### Task 5: Scalar Types - Error, Variant, Unknown, Dispatch, Empty

**Files:**
- Create: `TestCOMServer/TestCOMServer/Types/Error.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Variant.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Unknown.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Dispatch.cs`
- Create: `TestCOMServer/TestCOMServer/Types/Empty.cs`

- [ ] **Step 1: Create Types/Error.cs**

VT_ERROR holds an SCODE (HRESULT). Marshaled via `MarshalAs(UnmanagedType.Error)`.

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000017"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestError
    {
        [return: MarshalAs(UnmanagedType.Error)]
        int ErrorField
        {
            [return: MarshalAs(UnmanagedType.Error)]
            get;
            [param: MarshalAs(UnmanagedType.Error)]
            set;
        }

        int PutError([MarshalAs(UnmanagedType.Error)] int value);

        [return: MarshalAs(UnmanagedType.Error)]
        int GetError();

        [return: MarshalAs(UnmanagedType.Error)]
        int EchoError([MarshalAs(UnmanagedType.Error)] int value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000018"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Error"),
        ComSourceInterfaces(typeof(ICOMTestError))]
    public class COMTestError : ICOMTestError
    {
        protected int rawError;

        public COMTestError()
        {
            rawError = 0;
        }

        public int ErrorField
        {
            get { return rawError; }
            set { rawError = value; }
        }

        public int PutError(int value) { rawError = value; return 1; }
        public int GetError() { return rawError; }
        public int EchoError(int value) { return value; }
    }
}
```

- [ ] **Step 2: Create Types/Variant.cs**

VT_VARIANT is a generic VARIANT — `object` in .NET COM interop.

```csharp
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
```

- [ ] **Step 3: Create Types/Unknown.cs**

VT_UNKNOWN is an IUnknown pointer. Marshaled via `MarshalAs(UnmanagedType.IUnknown)`.

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001B"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestUnknown
    {
        [return: MarshalAs(UnmanagedType.IUnknown)]
        object UnknownField
        {
            [return: MarshalAs(UnmanagedType.IUnknown)]
            get;
            [param: MarshalAs(UnmanagedType.IUnknown)]
            set;
        }

        int PutUnknown([MarshalAs(UnmanagedType.IUnknown)] object value);

        [return: MarshalAs(UnmanagedType.IUnknown)]
        object GetUnknown();

        [return: MarshalAs(UnmanagedType.IUnknown)]
        object EchoUnknown([MarshalAs(UnmanagedType.IUnknown)] object value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001C"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Unknown"),
        ComSourceInterfaces(typeof(ICOMTestUnknown))]
    public class COMTestUnknown : ICOMTestUnknown
    {
        protected object rawUnknown;

        public COMTestUnknown()
        {
            rawUnknown = null;
        }

        public object UnknownField
        {
            get { return rawUnknown; }
            set { rawUnknown = value; }
        }

        public int PutUnknown(object value) { rawUnknown = value; return 1; }
        public object GetUnknown() { return rawUnknown; }
        public object EchoUnknown(object value) { return value; }
    }
}
```

- [ ] **Step 4: Create Types/Dispatch.cs**

VT_DISPATCH is an IDispatch pointer. Marshaled via `MarshalAs(UnmanagedType.IDispatch)`.

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001D"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestDispatch
    {
        [return: MarshalAs(UnmanagedType.IDispatch)]
        object DispatchField
        {
            [return: MarshalAs(UnmanagedType.IDispatch)]
            get;
            [param: MarshalAs(UnmanagedType.IDispatch)]
            set;
        }

        int PutDispatch([MarshalAs(UnmanagedType.IDispatch)] object value);

        [return: MarshalAs(UnmanagedType.IDispatch)]
        object GetDispatch();

        [return: MarshalAs(UnmanagedType.IDispatch)]
        object EchoDispatch([MarshalAs(UnmanagedType.IDispatch)] object value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001E"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Dispatch"),
        ComSourceInterfaces(typeof(ICOMTestDispatch))]
    public class COMTestDispatch : ICOMTestDispatch
    {
        protected object rawDispatch;

        public COMTestDispatch()
        {
            rawDispatch = null;
        }

        public object DispatchField
        {
            get { return rawDispatch; }
            set { rawDispatch = value; }
        }

        public int PutDispatch(object value) { rawDispatch = value; return 1; }
        public object GetDispatch() { return rawDispatch; }
        public object EchoDispatch(object value) { return value; }
    }
}
```

- [ ] **Step 5: Create Types/Empty.cs**

VT_EMPTY and VT_NULL testing. Methods that return null/empty variants and test null handling.

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-00000000001F"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestEmpty
    {
        object GetEmpty();
        object GetNull();
        bool IsEmpty(object value);
        bool IsNull(object value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-1111-1111-1111-000000000020"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Empty"),
        ComSourceInterfaces(typeof(ICOMTestEmpty))]
    public class COMTestEmpty : ICOMTestEmpty
    {
        public COMTestEmpty() { }

        public object GetEmpty() { return null; }
        public object GetNull() { return DBNull.Value; }
        public bool IsEmpty(object value) { return value == null; }
        public bool IsNull(object value) { return value is DBNull; }
    }
}
```

- [ ] **Step 6: Verify build compiles**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 7: Commit**

```bash
git add TestCOMServer/TestCOMServer/Types/Error.cs TestCOMServer/TestCOMServer/Types/Variant.cs TestCOMServer/TestCOMServer/Types/Unknown.cs TestCOMServer/TestCOMServer/Types/Dispatch.cs TestCOMServer/TestCOMServer/Types/Empty.cs
git commit -m "Add scalar types: Error, Variant, Unknown, Dispatch, Empty"
```

---

### Task 6: SafeArrays - Integer Types (Int8, Int16, Int32, Int64)

**Files:**
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Int8.cs`
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Int16.cs`
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Int32.cs`
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Int64.cs`

- [ ] **Step 1: Create SafeArrays/Int8.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000001"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayInt8
    {
        sbyte[] EchoInt8Array1D(sbyte[] value);
        sbyte[,] EchoInt8Array2D(sbyte[,] value);
        sbyte[,,] EchoInt8Array3D(sbyte[,,] value);
        byte[] EchoUInt8Array1D(byte[] value);
        byte[,] EchoUInt8Array2D(byte[,] value);
        byte[,,] EchoUInt8Array3D(byte[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000002"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Int8"),
        ComSourceInterfaces(typeof(ICOMSafeArrayInt8))]
    public class COMSafeArrayInt8 : ICOMSafeArrayInt8
    {
        public COMSafeArrayInt8() { }

        public sbyte[] EchoInt8Array1D(sbyte[] value) { return value; }
        public sbyte[,] EchoInt8Array2D(sbyte[,] value) { return value; }
        public sbyte[,,] EchoInt8Array3D(sbyte[,,] value) { return value; }
        public byte[] EchoUInt8Array1D(byte[] value) { return value; }
        public byte[,] EchoUInt8Array2D(byte[,] value) { return value; }
        public byte[,,] EchoUInt8Array3D(byte[,,] value) { return value; }
    }
}
```

- [ ] **Step 2: Create SafeArrays/Int16.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000003"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayInt16
    {
        short[] EchoInt16Array1D(short[] value);
        short[,] EchoInt16Array2D(short[,] value);
        short[,,] EchoInt16Array3D(short[,,] value);
        ushort[] EchoUInt16Array1D(ushort[] value);
        ushort[,] EchoUInt16Array2D(ushort[,] value);
        ushort[,,] EchoUInt16Array3D(ushort[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000004"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Int16"),
        ComSourceInterfaces(typeof(ICOMSafeArrayInt16))]
    public class COMSafeArrayInt16 : ICOMSafeArrayInt16
    {
        public COMSafeArrayInt16() { }

        public short[] EchoInt16Array1D(short[] value) { return value; }
        public short[,] EchoInt16Array2D(short[,] value) { return value; }
        public short[,,] EchoInt16Array3D(short[,,] value) { return value; }
        public ushort[] EchoUInt16Array1D(ushort[] value) { return value; }
        public ushort[,] EchoUInt16Array2D(ushort[,] value) { return value; }
        public ushort[,,] EchoUInt16Array3D(ushort[,,] value) { return value; }
    }
}
```

- [ ] **Step 3: Create SafeArrays/Int32.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000005"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayInt32
    {
        int[] EchoInt32Array1D(int[] value);
        int[,] EchoInt32Array2D(int[,] value);
        int[,,] EchoInt32Array3D(int[,,] value);
        uint[] EchoUInt32Array1D(uint[] value);
        uint[,] EchoUInt32Array2D(uint[,] value);
        uint[,,] EchoUInt32Array3D(uint[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000006"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Int32"),
        ComSourceInterfaces(typeof(ICOMSafeArrayInt32))]
    public class COMSafeArrayInt32 : ICOMSafeArrayInt32
    {
        public COMSafeArrayInt32() { }

        public int[] EchoInt32Array1D(int[] value) { return value; }
        public int[,] EchoInt32Array2D(int[,] value) { return value; }
        public int[,,] EchoInt32Array3D(int[,,] value) { return value; }
        public uint[] EchoUInt32Array1D(uint[] value) { return value; }
        public uint[,] EchoUInt32Array2D(uint[,] value) { return value; }
        public uint[,,] EchoUInt32Array3D(uint[,,] value) { return value; }
    }
}
```

- [ ] **Step 4: Create SafeArrays/Int64.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000007"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayInt64
    {
        long[] EchoInt64Array1D(long[] value);
        long[,] EchoInt64Array2D(long[,] value);
        long[,,] EchoInt64Array3D(long[,,] value);
        ulong[] EchoUInt64Array1D(ulong[] value);
        ulong[,] EchoUInt64Array2D(ulong[,] value);
        ulong[,,] EchoUInt64Array3D(ulong[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000008"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Int64"),
        ComSourceInterfaces(typeof(ICOMSafeArrayInt64))]
    public class COMSafeArrayInt64 : ICOMSafeArrayInt64
    {
        public COMSafeArrayInt64() { }

        public long[] EchoInt64Array1D(long[] value) { return value; }
        public long[,] EchoInt64Array2D(long[,] value) { return value; }
        public long[,,] EchoInt64Array3D(long[,,] value) { return value; }
        public ulong[] EchoUInt64Array1D(ulong[] value) { return value; }
        public ulong[,] EchoUInt64Array2D(ulong[,] value) { return value; }
        public ulong[,,] EchoUInt64Array3D(ulong[,,] value) { return value; }
    }
}
```

- [ ] **Step 5: Verify build compiles**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 6: Commit**

```bash
git add TestCOMServer/TestCOMServer/SafeArrays/Int8.cs TestCOMServer/TestCOMServer/SafeArrays/Int16.cs TestCOMServer/TestCOMServer/SafeArrays/Int32.cs TestCOMServer/TestCOMServer/SafeArrays/Int64.cs
git commit -m "Add SafeArray types: Int8, Int16, Int32, Int64 with 1D/2D/3D"
```

---

### Task 7: SafeArrays - Float32, Float64, String, Boolean

**Files:**
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Float32.cs`
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Float64.cs`
- Create: `TestCOMServer/TestCOMServer/SafeArrays/String.cs`
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Boolean.cs`

- [ ] **Step 1: Create SafeArrays/Float32.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000009"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayFloat32
    {
        float[] EchoFloat32Array1D(float[] value);
        float[,] EchoFloat32Array2D(float[,] value);
        float[,,] EchoFloat32Array3D(float[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000A"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Float32"),
        ComSourceInterfaces(typeof(ICOMSafeArrayFloat32))]
    public class COMSafeArrayFloat32 : ICOMSafeArrayFloat32
    {
        public COMSafeArrayFloat32() { }

        public float[] EchoFloat32Array1D(float[] value) { return value; }
        public float[,] EchoFloat32Array2D(float[,] value) { return value; }
        public float[,,] EchoFloat32Array3D(float[,,] value) { return value; }
    }
}
```

- [ ] **Step 2: Create SafeArrays/Float64.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000B"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayFloat64
    {
        double[] EchoFloat64Array1D(double[] value);
        double[,] EchoFloat64Array2D(double[,] value);
        double[,,] EchoFloat64Array3D(double[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000C"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Float64"),
        ComSourceInterfaces(typeof(ICOMSafeArrayFloat64))]
    public class COMSafeArrayFloat64 : ICOMSafeArrayFloat64
    {
        public COMSafeArrayFloat64() { }

        public double[] EchoFloat64Array1D(double[] value) { return value; }
        public double[,] EchoFloat64Array2D(double[,] value) { return value; }
        public double[,,] EchoFloat64Array3D(double[,,] value) { return value; }
    }
}
```

- [ ] **Step 3: Create SafeArrays/String.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000D"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayString
    {
        string[] EchoStringArray1D(string[] value);
        string[,] EchoStringArray2D(string[,] value);
        string[,,] EchoStringArray3D(string[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000E"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.String"),
        ComSourceInterfaces(typeof(ICOMSafeArrayString))]
    public class COMSafeArrayString : ICOMSafeArrayString
    {
        public COMSafeArrayString() { }

        public string[] EchoStringArray1D(string[] value) { return value; }
        public string[,] EchoStringArray2D(string[,] value) { return value; }
        public string[,,] EchoStringArray3D(string[,,] value) { return value; }
    }
}
```

- [ ] **Step 4: Create SafeArrays/Boolean.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-00000000000F"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayBoolean
    {
        bool[] EchoBooleanArray1D(bool[] value);
        bool[,] EchoBooleanArray2D(bool[,] value);
        bool[,,] EchoBooleanArray3D(bool[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000010"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Boolean"),
        ComSourceInterfaces(typeof(ICOMSafeArrayBoolean))]
    public class COMSafeArrayBoolean : ICOMSafeArrayBoolean
    {
        public COMSafeArrayBoolean() { }

        public bool[] EchoBooleanArray1D(bool[] value) { return value; }
        public bool[,] EchoBooleanArray2D(bool[,] value) { return value; }
        public bool[,,] EchoBooleanArray3D(bool[,,] value) { return value; }
    }
}
```

- [ ] **Step 5: Verify build compiles**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 6: Commit**

```bash
git add TestCOMServer/TestCOMServer/SafeArrays/Float32.cs TestCOMServer/TestCOMServer/SafeArrays/Float64.cs TestCOMServer/TestCOMServer/SafeArrays/String.cs TestCOMServer/TestCOMServer/SafeArrays/Boolean.cs
git commit -m "Add SafeArray types: Float32, Float64, String, Boolean with 1D/2D/3D"
```

---

### Task 8: SafeArrays - Currency, Date, Decimal, Variant

**Files:**
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Currency.cs`
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Date.cs`
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Decimal.cs`
- Create: `TestCOMServer/TestCOMServer/SafeArrays/Variant.cs`

- [ ] **Step 1: Create SafeArrays/Currency.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000011"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayCurrency
    {
        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)]
        decimal[] EchoCurrencyArray1D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)] decimal[] value);

        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)]
        decimal[,] EchoCurrencyArray2D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)] decimal[,] value);

        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)]
        decimal[,,] EchoCurrencyArray3D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_CY)] decimal[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000012"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Currency"),
        ComSourceInterfaces(typeof(ICOMSafeArrayCurrency))]
    public class COMSafeArrayCurrency : ICOMSafeArrayCurrency
    {
        public COMSafeArrayCurrency() { }

        public decimal[] EchoCurrencyArray1D(decimal[] value) { return value; }
        public decimal[,] EchoCurrencyArray2D(decimal[,] value) { return value; }
        public decimal[,,] EchoCurrencyArray3D(decimal[,,] value) { return value; }
    }
}
```

- [ ] **Step 2: Create SafeArrays/Date.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000013"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayDate
    {
        DateTime[] EchoDateArray1D(DateTime[] value);
        DateTime[,] EchoDateArray2D(DateTime[,] value);
        DateTime[,,] EchoDateArray3D(DateTime[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000014"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Date"),
        ComSourceInterfaces(typeof(ICOMSafeArrayDate))]
    public class COMSafeArrayDate : ICOMSafeArrayDate
    {
        public COMSafeArrayDate() { }

        public DateTime[] EchoDateArray1D(DateTime[] value) { return value; }
        public DateTime[,] EchoDateArray2D(DateTime[,] value) { return value; }
        public DateTime[,,] EchoDateArray3D(DateTime[,,] value) { return value; }
    }
}
```

- [ ] **Step 3: Create SafeArrays/Decimal.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000015"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayDecimal
    {
        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)]
        decimal[] EchoDecimalArray1D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)] decimal[] value);

        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)]
        decimal[,] EchoDecimalArray2D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)] decimal[,] value);

        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)]
        decimal[,,] EchoDecimalArray3D(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DECIMAL)] decimal[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000016"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Decimal"),
        ComSourceInterfaces(typeof(ICOMSafeArrayDecimal))]
    public class COMSafeArrayDecimal : ICOMSafeArrayDecimal
    {
        public COMSafeArrayDecimal() { }

        public decimal[] EchoDecimalArray1D(decimal[] value) { return value; }
        public decimal[,] EchoDecimalArray2D(decimal[,] value) { return value; }
        public decimal[,,] EchoDecimalArray3D(decimal[,,] value) { return value; }
    }
}
```

- [ ] **Step 4: Create SafeArrays/Variant.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000017"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMSafeArrayVariant
    {
        object[] EchoVariantArray1D(object[] value);
        object[,] EchoVariantArray2D(object[,] value);
        object[,,] EchoVariantArray3D(object[,,] value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-2222-2222-2222-000000000018"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.SafeArray.Variant"),
        ComSourceInterfaces(typeof(ICOMSafeArrayVariant))]
    public class COMSafeArrayVariant : ICOMSafeArrayVariant
    {
        public COMSafeArrayVariant() { }

        public object[] EchoVariantArray1D(object[] value) { return value; }
        public object[,] EchoVariantArray2D(object[,] value) { return value; }
        public object[,,] EchoVariantArray3D(object[,,] value) { return value; }
    }
}
```

- [ ] **Step 5: Verify build compiles**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 6: Commit**

```bash
git add TestCOMServer/TestCOMServer/SafeArrays/Currency.cs TestCOMServer/TestCOMServer/SafeArrays/Date.cs TestCOMServer/TestCOMServer/SafeArrays/Decimal.cs TestCOMServer/TestCOMServer/SafeArrays/Variant.cs
git commit -m "Add SafeArray types: Currency, Date, Decimal, Variant with 1D/2D/3D"
```

---

### Task 9: Interface Patterns - DualInterface, UnknownOnly, DispatchOnly

**Files:**
- Create: `TestCOMServer/TestCOMServer/Interfaces/DualInterface.cs`
- Create: `TestCOMServer/TestCOMServer/Interfaces/UnknownOnly.cs`
- Create: `TestCOMServer/TestCOMServer/Interfaces/DispatchOnly.cs`

- [ ] **Step 1: Create Interfaces/DualInterface.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000001"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestDualInterface
    {
        int GetValue();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000002"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.Dual"),
        ComSourceInterfaces(typeof(ICOMTestDualInterface))]
    public class COMTestDualInterface : ICOMTestDualInterface
    {
        public COMTestDualInterface() { }

        public int GetValue() { return 42; }
    }
}
```

- [ ] **Step 2: Create Interfaces/UnknownOnly.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000003"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ICOMTestUnknownOnly
    {
        int GetValue();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000004"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.UnknownOnly"),
        ComSourceInterfaces(typeof(ICOMTestUnknownOnly))]
    public class COMTestUnknownOnly : ICOMTestUnknownOnly
    {
        public COMTestUnknownOnly() { }

        public int GetValue() { return 84; }
    }
}
```

- [ ] **Step 3: Create Interfaces/DispatchOnly.cs**

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000005"),
        InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICOMTestDispatchOnly
    {
        int DispatchGetValue();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000006"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.DispatchOnly"),
        ComSourceInterfaces(typeof(ICOMTestDispatchOnly))]
    public class COMTestDispatchOnly : ICOMTestDispatchOnly
    {
        public COMTestDispatchOnly() { }

        public int DispatchGetValue() { return 126; }
    }
}
```

- [ ] **Step 4: Verify build compiles**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 5: Commit**

```bash
git add TestCOMServer/TestCOMServer/Interfaces/DualInterface.cs TestCOMServer/TestCOMServer/Interfaces/UnknownOnly.cs TestCOMServer/TestCOMServer/Interfaces/DispatchOnly.cs
git commit -m "Add interface patterns: Dual, IUnknown-only, IDispatch-only"
```

---

### Task 10: Interface Patterns - MultipleInterfaces, InheritedInterface

**Files:**
- Create: `TestCOMServer/TestCOMServer/Interfaces/MultipleInterfaces.cs`
- Create: `TestCOMServer/TestCOMServer/Interfaces/InheritedInterface.cs`

- [ ] **Step 1: Create Interfaces/MultipleInterfaces.cs**

Three separate interfaces implemented by a single class. Each returns a different constant.

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000007"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestMultiA
    {
        int GetValueA();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000008"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestMultiB
    {
        int GetValueB();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000009"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestMultiC
    {
        int GetValueC();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000A"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.Multiple"),
        ComSourceInterfaces(typeof(ICOMTestMultiA))]
    public class COMTestMultipleInterfaces : ICOMTestMultiA, ICOMTestMultiB, ICOMTestMultiC
    {
        public COMTestMultipleInterfaces() { }

        public int GetValueA() { return 10; }
        public int GetValueB() { return 20; }
        public int GetValueC() { return 30; }
    }
}
```

- [ ] **Step 2: Create Interfaces/InheritedInterface.cs**

Interface B inherits from interface A. Class implements B (and therefore A). Each level has its own constant.

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000B"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestBaseInterface
    {
        int GetBaseValue();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000C"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestDerivedInterface : ICOMTestBaseInterface
    {
        int GetDerivedValue();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000D"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.Inherited"),
        ComSourceInterfaces(typeof(ICOMTestDerivedInterface))]
    public class COMTestInheritedInterface : ICOMTestDerivedInterface
    {
        public COMTestInheritedInterface() { }

        public int GetBaseValue() { return 100; }
        public int GetDerivedValue() { return 200; }
    }
}
```

- [ ] **Step 3: Verify build compiles**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 4: Commit**

```bash
git add TestCOMServer/TestCOMServer/Interfaces/MultipleInterfaces.cs TestCOMServer/TestCOMServer/Interfaces/InheritedInterface.cs
git commit -m "Add interface patterns: Multiple interfaces, inherited interface"
```

---

### Task 11: Interface Patterns - ConnectionPoint

**Files:**
- Create: `TestCOMServer/TestCOMServer/Interfaces/ConnectionPoint.cs`

- [ ] **Step 1: Create Interfaces/ConnectionPoint.cs**

Implements COM connection points for event sourcing. Defines an event source interface and a class that fires events.

```csharp
using System;
using System.Runtime.InteropServices;

namespace TestCOMServer
{
    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000E"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMTestConnectionPoint
    {
        int QueryValue();
        void FireEvent();
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-00000000000F"),
        InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICOMTestConnectionPointEvents
    {
        [DispId(1)]
        void OnEvent(int value);
    }

    [ComVisible(true),
        Guid("A1B2C3D4-3333-3333-3333-000000000010"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("TestCOM.Interface.ConnectionPoint"),
        ComSourceInterfaces(typeof(ICOMTestConnectionPointEvents))]
    public class COMTestConnectionPoint : ICOMTestConnectionPoint
    {
        public delegate void OnEventHandler(int value);
        public event OnEventHandler OnEvent;

        public COMTestConnectionPoint() { }

        public int QueryValue() { return 999; }

        public void FireEvent()
        {
            OnEvent?.Invoke(999);
        }
    }
}
```

- [ ] **Step 2: Verify build compiles**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 3: Commit**

```bash
git add TestCOMServer/TestCOMServer/Interfaces/ConnectionPoint.cs
git commit -m "Add interface pattern: ConnectionPoint with COM event sourcing"
```

---

### Task 12: Update ChangeLog and Final Verification

**Files:**
- Modify: `ChangeLog.md`

- [ ] **Step 1: Update ChangeLog.md**

Add a new version entry at the top of the file, before the existing `# Version 1.0.2` entry:

```markdown
# Version 2.0.0

 1. Migrate from .NET Framework 4.5.2 to .NET 5.0 (SDK-style project).
 1. Reorganize codebase into Types/, SafeArrays/, and Interfaces/ directories.
 1. One COM class per type for isolated go-ole testing.
 1. New ProgID convention: `TestCOM.<Type>`, `TestCOM.SafeArray.<Type>`, `TestCOM.Interface.<Pattern>`.
 1. Scalar VARIANT types:
    1. Existing: `Int8/UInt8`, `Int16/UInt16`, `Int32/UInt32`, `Int64/UInt64`, `Float32`, `Float64`, `String`, `Boolean`.
    1. New: `Currency` (VT_CY), `Date` (VT_DATE), `Decimal` (VT_DECIMAL), `Error` (VT_ERROR), `Variant` (VT_VARIANT), `Unknown` (VT_UNKNOWN), `Dispatch` (VT_DISPATCH), `Empty/Null` (VT_EMPTY/VT_NULL).
 1. SafeArray support with 1D, 2D, and 3D arrays for: Int8/UInt8, Int16/UInt16, Int32/UInt32, Int64/UInt64, Float32, Float64, String, Boolean, Currency, Date, Decimal, Variant.
 1. COM interface pattern test classes:
    1. Dual interface (IDispatch + vtable).
    1. IUnknown-only interface (vtable only).
    1. IDispatch-only interface (late binding only).
    1. Multiple interfaces on single class (QueryInterface testing).
    1. Inherited interfaces (base + derived).
    1. Connection points (COM event sourcing).
 1. Delete old `COMTestDispatch.cs` and `COMSafeArray.cs`.
```

- [ ] **Step 2: Verify full build from clean state**

```bash
cd /Users/jacobsantos/repos/respite-studios/test-com-server/TestCOMServer
dotnet clean
dotnet build
```

Expected: Build succeeded with 0 errors.

- [ ] **Step 3: Commit**

```bash
git add ChangeLog.md
git commit -m "Update ChangeLog for v2.0.0 release"
```
