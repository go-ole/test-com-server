# COM Server Full VARIANT Type & Interface Support

## Purpose

Rebuild the .NET COM test server to support all VARIANT types, multi-dimensional SAFEARRAYs for every supported array element type, and all COM interface patterns. This serves as:

1. **go-ole test coverage** — comprehensive type and interface marshaling validation
2. **COM interop reference** — demonstrates every VARIANT type and interface pattern in .NET
3. **Interface verification** — tests QueryInterface, IUnknown, IDispatch, dual interfaces, connection points, and interface inheritance

## Scope

- All representable VARIANT types as scalar get/set/echo COM classes
- 1D, 2D, and 3D SAFEARRAY echo classes for every supported array element type
- COM interface pattern test classes (dual, IUnknown-only, IDispatch-only, multiple interfaces, inheritance, connection points)
- Delete existing `COMTestDispatch.cs` and `COMSafeArray.cs`
- Organized under `Types/`, `SafeArrays/`, `Interfaces/` directories

## Architecture

### Design Principles

- **One COM class per type** — maximum isolation for go-ole testing; a failing test points directly at the problem
- **Consistent pattern** — every type class follows the same interface+class template (get/set property, Put/Get/Echo methods)
- **Simple interface tests** — interface pattern classes return hardcoded constants, not full type matrices
- **ProgID convention** — `TestCOM.<Type>`, `TestCOM.SafeArray.<Type>`, `TestCOM.Interface.<Pattern>`
- **Fresh GUIDs** — every new interface and class gets a new GUID; no reuse from deleted classes

### Directory Layout

```
TestCOMServer/TestCOMServer/
  Types/
    Int8.cs, Int16.cs, Int32.cs, Int64.cs,
    Float32.cs, Float64.cs, String.cs, Boolean.cs,
    Currency.cs, Date.cs, Decimal.cs, Error.cs,
    Variant.cs, Unknown.cs, Dispatch.cs, Empty.cs
  SafeArrays/
    Int8.cs, Int16.cs, Int32.cs, Int64.cs,
    Float32.cs, Float64.cs, String.cs, Boolean.cs,
    Currency.cs, Date.cs, Decimal.cs, Variant.cs
  Interfaces/
    DualInterface.cs, UnknownOnly.cs, DispatchOnly.cs,
    MultipleInterfaces.cs, InheritedInterface.cs, ConnectionPoint.cs
```

## Section 1: Scalar Types (`Types/`)

Each VARIANT type gets one interface + one class. The pattern for each:

```csharp
[ComVisible(true), Guid("..."), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface ICOMTest<Type> {
    <type> <Type>Field { get; set; }
    int Put<Type>(<type> value);
    <type> Get<Type>();
    <type> Echo<Type>(<type> value);
}

[ComVisible(true), Guid("..."), ClassInterface(ClassInterfaceType.None), ProgId("TestCOM.<Type>")]
public class COMTest<Type> : ICOMTest<Type> {
    // protected raw<Type> backing field
    // public no-arg constructor initializing to default
    // Put returns 1 (or string length for strings)
    // Echo returns input unchanged
}
```

### Type Files

| File | ProgID | VARIANT Type | .NET Type |
|------|--------|-------------|-----------|
| `Types/Int8.cs` | TestCOM.Int8 | VT_I1 / VT_UI1 | sbyte / byte (both in one class) |
| `Types/Int16.cs` | TestCOM.Int16 | VT_I2 / VT_UI2 | short / ushort |
| `Types/Int32.cs` | TestCOM.Int32 | VT_I4 / VT_UI4 | int / uint |
| `Types/Int64.cs` | TestCOM.Int64 | VT_I8 / VT_UI8 | long / ulong |
| `Types/Float32.cs` | TestCOM.Float32 | VT_R4 | float |
| `Types/Float64.cs` | TestCOM.Float64 | VT_R8 | double |
| `Types/String.cs` | TestCOM.String | VT_BSTR | string |
| `Types/Boolean.cs` | TestCOM.Boolean | VT_BOOL | bool |
| `Types/Currency.cs` | TestCOM.Currency | VT_CY | long (scaled by 10000, OLE currency) |
| `Types/Date.cs` | TestCOM.Date | VT_DATE | DateTime |
| `Types/Decimal.cs` | TestCOM.Decimal | VT_DECIMAL | decimal |
| `Types/Error.cs` | TestCOM.Error | VT_ERROR | int (SCODE/HRESULT) |
| `Types/Variant.cs` | TestCOM.Variant | VT_VARIANT | object |
| `Types/Unknown.cs` | TestCOM.Unknown | VT_UNKNOWN | object (IUnknown pointer) |
| `Types/Dispatch.cs` | TestCOM.Dispatch | VT_DISPATCH | object (IDispatch pointer) |
| `Types/Empty.cs` | TestCOM.Empty | VT_EMPTY / VT_NULL | Tests returning null/empty variants |

Integer files (Int8, Int16, Int32, Int64) each handle both signed and unsigned variants in a single interface and class — matching the existing convention with separate properties and Put/Get/Echo methods for each signedness.

## Section 2: SafeArrays (`SafeArrays/`)

Each supported SAFEARRAY element type gets one interface + one class with echo methods for 1D, 2D, and 3D arrays:

```csharp
[ComVisible(true), Guid("..."), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface ICOMSafeArray<Type> {
    <type>[] Echo<Type>Array1D(<type>[] value);
    <type>[,] Echo<Type>Array2D(<type>[,] value);
    <type>[,,] Echo<Type>Array3D(<type>[,,] value);
}

[ComVisible(true), Guid("..."), ClassInterface(ClassInterfaceType.None), ProgId("TestCOM.SafeArray.<Type>")]
public class COMSafeArray<Type> : ICOMSafeArray<Type> {
    // All methods are pure echo — return input unchanged
}
```

Integer files handle both signed and unsigned (6 methods each: signed 1D/2D/3D + unsigned 1D/2D/3D).

### SafeArray Files

| File | ProgID | Element Type |
|------|--------|-------------|
| `SafeArrays/Int8.cs` | TestCOM.SafeArray.Int8 | VT_I1 / VT_UI1 |
| `SafeArrays/Int16.cs` | TestCOM.SafeArray.Int16 | VT_I2 / VT_UI2 |
| `SafeArrays/Int32.cs` | TestCOM.SafeArray.Int32 | VT_I4 / VT_UI4 |
| `SafeArrays/Int64.cs` | TestCOM.SafeArray.Int64 | VT_I8 / VT_UI8 |
| `SafeArrays/Float32.cs` | TestCOM.SafeArray.Float32 | VT_R4 |
| `SafeArrays/Float64.cs` | TestCOM.SafeArray.Float64 | VT_R8 |
| `SafeArrays/String.cs` | TestCOM.SafeArray.String | VT_BSTR |
| `SafeArrays/Boolean.cs` | TestCOM.SafeArray.Boolean | VT_BOOL |
| `SafeArrays/Currency.cs` | TestCOM.SafeArray.Currency | VT_CY |
| `SafeArrays/Date.cs` | TestCOM.SafeArray.Date | VT_DATE |
| `SafeArrays/Decimal.cs` | TestCOM.SafeArray.Decimal | VT_DECIMAL |
| `SafeArrays/Variant.cs` | TestCOM.SafeArray.Variant | VT_VARIANT |

**Excluded from SafeArray:** VT_ERROR, VT_UNKNOWN, VT_DISPATCH, VT_EMPTY/VT_NULL — these are not standard SAFEARRAY element types.

## Section 3: Interface Patterns (`Interfaces/`)

Each class exposes a simple hardcoded constant value to verify interface mechanics from go-ole. The focus is testing QueryInterface behavior, not data types.

### Interface Files

| File | ProgID | What It Tests |
|------|--------|--------------|
| `Interfaces/DualInterface.cs` | TestCOM.Interface.Dual | `InterfaceIsDual` — supports both IDispatch and vtable binding. `GetValue()` returns constant. |
| `Interfaces/UnknownOnly.cs` | TestCOM.Interface.UnknownOnly | `InterfaceIsIUnknown` — vtable only, no IDispatch. `GetValue()` returns constant. |
| `Interfaces/DispatchOnly.cs` | TestCOM.Interface.DispatchOnly | `InterfaceIsIDispatch` — late binding only. `GetValue()` returns constant. |
| `Interfaces/MultipleInterfaces.cs` | TestCOM.Interface.Multiple | Single class implementing 3 separate interfaces (each with its own GUID and a different constant). Tests QueryInterface across interfaces. |
| `Interfaces/InheritedInterface.cs` | TestCOM.Interface.Inherited | Interface B inherits from interface A. Class implements B. Tests that QueryInterface resolves both A and B. Each level has its own constant. |
| `Interfaces/ConnectionPoint.cs` | TestCOM.Interface.ConnectionPoint | Implements `IConnectionPoint` / `IConnectionPointContainer` for COM event sourcing. Fires a simple event with a constant value. |

Each constant is a different value so go-ole tests can verify they received the correct interface/object.

## Section 4: Project & Build Changes

- **Delete** `COMTestDispatch.cs` and `COMSafeArray.cs`
- **Update** `.csproj` to include all new files under `Types\`, `SafeArrays\`, `Interfaces\`
- **Namespace:** `TestCOMServer` (unchanged)
- **Target framework:** .NET Framework 4.5.2 (unchanged)
- **Assembly version:** Bump to 2.0.0.0 (breaking change — old classes removed)
- **Assembly GUID:** Keep existing `ac11d114-ebee-4d2f-b968-127f155e4f6e`
- **Build configurations:** Keep AnyCPU, x86, x64 with existing post-build tlbexp
- **GUIDs:** Every new interface and class gets a freshly generated GUID. No reuse from deleted classes.
