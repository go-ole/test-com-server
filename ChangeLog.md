# Version 2.0.0

 1. Migrate from .NET Framework 4.5.2 to .NET 5.0 (SDK-style project).
 1. Reorganize codebase into Types/, SafeArrays/, and Interfaces/ directories.
 1. One COM class per type for isolated go-ole testing.
 1. New ProgID convention: `TestCOM.<Type>`, `TestCOM.SafeArray.<Type>`, `TestCOM.Interface.<Pattern>`.
 1. Scalar VARIANT types:
    1. Existing: `Int8/UInt8`, `Int16/UInt16`, `Int32/UInt32`, `Int64/UInt64`, `Float32`, `Float64`, `String`, `Boolean`.
    1. New: `Currency` (VT_CY), `Date` (VT_DATE), `Decimal` (VT_DECIMAL), `Error` (VT_ERROR), `Variant` (VT_VARIANT), `Unknown` (VT_UNKNOWN), `Dispatch` (VT_DISPATCH), `Empty/Null` (VT_EMPTY/VT_NULL), `Clsid` (VT_CLSID), `HResult` (VT_HRESULT), `FileTime` (VT_FILETIME), `Blob` (VT_BLOB as SAFEARRAY VT_UI1), `Ptr` (VT_PTR via IUnknown-only), `Stream` (VT_STREAM with IStream implementation).
 1. SafeArray support with 1D, 2D, and 3D arrays for: Int8/UInt8, Int16/UInt16, Int32/UInt32, Int64/UInt64, Float32, Float64, String, Boolean, Currency, Date, Decimal, Variant.
 1. COM interface pattern test classes:
    1. Dual interface (IDispatch + vtable).
    1. IUnknown-only interface (vtable only).
    1. IDispatch-only interface (late binding only).
    1. Multiple interfaces on single class (QueryInterface testing).
    1. Inherited interfaces (base + derived).
    1. Connection points (COM event sourcing).
 1. Delete old `COMTestDispatch.cs` and `COMSafeArray.cs`.

# Version 1.0.2 (2015-11-02)

 1. Fix interface COM visibility by copying properties and methods to interface.
 1. Make constructors public.
 1. Add boolean implementation to interface.

# Version 1.0.1 (2015-10-30)

 1. Add safe array server implementation.
 1. Make methods public.

# Version 1.0.0 (2015-08-05)

 1. Support for accessing, mutating and echoing types:
    1. `Int8`, `UInt8`, `Int16`, `UInt16`, `Int32`, `UInt32`, `Int64`, `UInt64`
    1. `Float32`, `Float64`
    1. Unicode String.
 1. COM Interfaces (Name - GUID):
    1. `ICOMTestString` (E0133EB4-C36F-469A-9D3D-C66B84BE19ED)
    1. `ICOMTestInt8` (BEB06610-EB84-4155-AF58-E2BFF53608B4)
    1. `ICOMTestInt16` (DAA3F9FA-761E-4976-A860-8364CE55F6FC)
    1. `ICOMTestInt32` (E3DEDEE7-38A2-4540-91D1-2EEF1D8891B0)
    1. `ICOMTestInt64` (8D437CBC-B3ED-485C-BC32-C336432A1623)
    1. `ICOMTestFloat` (BF1ED004-EA02-456A-AA55-2AC8AC6B054C)
    1. `ICOMTestDouble` (BF908A81-8687-4E93-999F-D86FAB284BA0)
    1. `ICOMTestBoolean` (D530E7A6-4EE8-40D1-8931-3D63B8605001)
    1. `ICOMTestTypes` (CCA8D7AE-91C0-4277-A8B3-FF4EDF28D3C0)
    1. `ICOMEchoTestObject` (6485B1EF-D780-4834-A4FE-1EBB51746CA3)
 1. COM Classes (Name - GUID)
    1. `COMEchoTestObject` (3C24506A-AE9E-4D50-9157-EF317281F1B0)
    1. `COMTestScalarClass` (865B85C5-0334-4AC6-9EF6-AACEC8FC5E86)
 1. Set x64 as default platform.
 1. Add Appveyor CI build with badge.
 1. Fix builds for AnyCPU and x64 platforms.
 1. Only make specific objects COM visible.
 1. Rename Float/Double to Float32/Float64.
 1. Create TLB assembly.
