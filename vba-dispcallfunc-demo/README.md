# DispCallFunc bridge demo (VBA / x64 Excel on Windows)

A minimal, benign demonstration of how `DispCallFunc` (exported by `OleAut32.dll`) lets VBA invoke a raw function pointer with arbitrary argument types on x64 Windows. Both demos call ordinary, documented Win32 exports — there is no syscall-number resolution, no RWX allocation, and no runtime-built code in this folder.

## Why `DispCallFunc` at all on x64

VBA's `Declare PtrSafe Function ... Lib "..."` only lets you call a **named export** of a DLL — VBA emits the call site at compile time against that import. It has no language construct for "call this raw function-pointer value with these arguments." On x86 you could fake it with tricks (function-pointer assignment via `CallWindowProc`, etc.), but on **x64 Windows there is exactly one calling convention** (Microsoft x64: first four integer/pointer args in `RCX/RDX/R8/R9`, rest spilled to the stack with a 32-byte shadow area, 16-byte stack alignment, return in `RAX`), and VBA gives you no way to construct that frame yourself.

`DispCallFunc` was designed for COM late binding — invoking an arbitrary slot in a vtable with arguments the caller only knows at runtime. Its signature lets you pass:

- a `pvInstance` (the COM object, or `NULL`),
- an `oVft` (vtable offset, **or — when `pvInstance` is `NULL` — a raw function pointer**),
- an array of `VARTYPE`s and an array of pointers to `VARIANTARG`s describing the arguments.

The `OleAut32` implementation marshals those VARIANTs into the proper x64 frame, executes the `call`, and writes the return value back into a VARIANT. So `DispCallFunc` is effectively a **calling-convention bridge**: it gives a high-level scripting host the ability to call any address with any signature, using the platform's native ABI. That's its legitimate purpose, and it's why it's the standard mechanism whenever VBA needs to invoke a function pointer it computed at runtime.

## What "SSN resolution" means (conceptually)

A user-mode program normally enters the kernel through stub functions exported by `ntdll.dll` — e.g., `NtAllocateVirtualMemory`, `NtProtectVirtualMemory`. Each stub is tiny: it loads a per-function integer called the **System Service Number (SSN)** into `EAX` and executes a `syscall` instruction, which transitions to ring 0. The kernel uses the SSN as an index into the System Service Descriptor Table (SSDT) to dispatch to the actual handler.

The SSN is **not a stable ABI**. Microsoft renumbers entries between Windows builds (and sometimes between security updates), so any code that wants to issue a `syscall` instruction directly — instead of `call ntdll!NtFoo` — has to figure out the correct SSN at runtime. "SSN resolution" is that figure-it-out step. Conceptually, the techniques fall into a few buckets:

- **Read the in-memory stub.** Find `ntdll!NtFoo` (e.g., parse the export table), then inspect the first few bytes; an unhooked stub starts with `mov eax, imm32` and the immediate is the SSN. (This is the original "Hell's Gate" idea.)
- **Infer from neighbours when the stub is hooked.** EDRs often replace the first bytes of the stub with a `jmp` into their own code, destroying the immediate. Because syscall numbers are roughly monotonic in export order, you can read SSNs from nearby unhooked stubs and adjust by index. (This is the "Halo's Gate" / "Tartarus' Gate" family.)
- **Static tables.** Hard-code SSNs per build number. Brittle but trivial.

The whole point of doing this in offensive tooling is that EDR user-mode hooks live in `ntdll`'s exported stubs; once you have the SSN, you can issue the `syscall` instruction yourself and never touch the hook. That's the detection-evasion property — and the reason this folder does **not** include that step.

## What the stages of a (hypothetical) chain are doing, in one line each

1. **Get a function-pointer-callable buffer.** Somewhere to host a tiny syscall stub.
2. **Lay down a stub.** A few bytes implementing the Microsoft x64 syscall convention (`mov r10, rcx; mov eax, <SSN>; syscall; ret`) — the `r10` move is needed because the `syscall` instruction itself clobbers `rcx`.
3. **Resolve the SSN** for the target `Nt*` using one of the techniques above.
4. **Invoke the stub via `DispCallFunc`,** because step 1's buffer isn't an exported DLL function — VBA can't call it with `Declare`, so the calling-convention bridge is mandatory.

The demo in this folder exercises only step 4's mechanism (`DispCallFunc` against a real, exported API) so you can see how the bridge works without any of the offensive pieces.

## What's in this folder

- [DispCallFuncDemo.bas](DispCallFuncDemo.bas) — VBA module with two procedures:
  - `Demo_GetTickCount` — zero-arg, 32-bit return. The minimal bridge call.
  - `Demo_MessageBoxW` — four args (two `LPCWSTR`, an `HWND`, a `UINT`), `int` return. This is the case where `DispCallFunc` actually has to lay out a real x64 frame (RCX/RDX/R8/R9 + shadow space) on your behalf.

## Running it

Tested with x64 Excel on Windows 11.

1. Open Excel and press `Alt+F11` to open the VBA editor.
2. `File → Import File…` and select `DispCallFuncDemo.bas`.
3. Open the Immediate window with `Ctrl+G`.
4. Type `RunAllDemos` and press Enter.

`GetTickCount`'s return value prints in the Immediate window. `MessageBoxW` pops a dialog and then prints the button code on dismissal.

## Notes on the code

- `pvInstance = 0` tells `DispCallFunc` to treat `oVft` as a raw function pointer instead of a vtable offset.
- `CC_STDCALL = 4` is the only `CALLCONV` value that matters on x64 — there's only one calling convention, and `OleAut32` expects this constant.
- `VT_I8` is used for everything pointer-sized (`HWND`, `LPCWSTR`) because on x64 those are 64-bit. `VT_I4` is used for true 32-bit ints like `UINT`.
- The `VARIANT`s holding the arguments must store values whose internal type matches the corresponding `VARTYPE` entry; that's why the demo uses `CLngLng` for the 64-bit slots and `CLng` for the 32-bit one.
- Even when `cActuals = 0`, `DispCallFunc` wants addressable storage for `prgvt`/`prgpvarg`; the `GetTickCount` demo passes dummy locals for that reason.
