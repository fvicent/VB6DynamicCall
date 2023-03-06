#include "pch.h"

#define DLL_EXPORT __declspec(dllexport)

/* Supported VB6 types. */
#define VB6_TYPE_BYTE 0
#define VB6_TYPE_DOUBLE 1
#define VB6_TYPE_INTEGER 2
#define VB6_TYPE_SINGLE 3
#define VB6_TYPE_STRING 4

/* These arrays hold the arguments pushed by the Attach* functions. */
BYTE *vb6ByteArgs;
unsigned int vb6ByteArgsCount;
WORD *vb6IntegerArgs;
unsigned int vb6IntegerArgsCount;
double *vb6DoubleArgs;
unsigned int vb6DoubleArgsCount;
BSTR *vb6StringArgs;
unsigned int vb6StringArgsCount;
float *vb6SingleArgs;
unsigned int vb6SingleArgsCount;

/* The type of each attached argument is stored in this array. */
int* attachedArgsTypes;

int GetAttachedArgsCount();

/* Internal functions */

void AddArgumentType(int type)
{
    int attachedArgsCount = GetAttachedArgsCount();
    int* originalAttachedArgsTypes = attachedArgsTypes;
    attachedArgsTypes = (int*)realloc(attachedArgsTypes, sizeof(int) * attachedArgsCount);
    if (attachedArgsTypes == NULL)
    {
        attachedArgsTypes = originalAttachedArgsTypes;
        OutputDebugString(L"Could not reallocate memory for attached args.");
        return;
    }
    attachedArgsTypes[attachedArgsCount - 1] = type;
}

int GetAttachedArgsCount()
{
    return (
        vb6ByteArgsCount + vb6DoubleArgsCount + vb6IntegerArgsCount
        + vb6SingleArgsCount + vb6StringArgsCount
    );
}

void ResetArgs()
{
    vb6ByteArgs = NULL;
    vb6DoubleArgs = NULL;
    vb6IntegerArgs = NULL;
    vb6SingleArgs = NULL;
    vb6StringArgs = NULL;
    vb6ByteArgsCount = 0;
    vb6DoubleArgsCount = 0;
    vb6IntegerArgsCount = 0;
    vb6SingleArgsCount = 0;
    vb6StringArgsCount = 0;
}


/* API functions */

DLL_EXPORT void WINAPI AttachByte(BYTE value)
{
    if (vb6ByteArgsCount == 0)
    {
        vb6ByteArgs = malloc(sizeof(BYTE));
        if (vb6ByteArgs == NULL)
        {
            OutputDebugString(L"Could not allocate memory for byte args.");
            return;
        }
        vb6ByteArgsCount++;
    }
    else
    {
        BYTE* originalByteArgs = vb6ByteArgs;
        vb6ByteArgs = realloc(vb6ByteArgs, sizeof(BYTE) * ++vb6ByteArgsCount);
        if (vb6ByteArgs == NULL)
        {
            vb6ByteArgs = originalByteArgs;
            OutputDebugString(L"Could not reallocate memory for byte args.");
            return;
        }
    }
    vb6ByteArgs[vb6ByteArgsCount - 1] = value;
    AddArgumentType(VB6_TYPE_BYTE);
    OutputDebugString(L"Byte argument attached.");
}

DLL_EXPORT void WINAPI AttachDouble(double value)
{
    if (vb6DoubleArgsCount == 0)
    {
        vb6DoubleArgs = malloc(sizeof(double));
        if (vb6DoubleArgs == NULL)
        {
            OutputDebugString(L"Could not allocate memory for double args.");
            return;
        }
        vb6DoubleArgsCount++;
    }
    else
    {
        double* originalDoubleArgs = vb6DoubleArgs;
        vb6DoubleArgs = realloc(vb6DoubleArgs, sizeof(double) * ++vb6DoubleArgsCount);
        if (vb6DoubleArgs == NULL)
        {
            vb6DoubleArgs = originalDoubleArgs;
            OutputDebugString(L"Could not reallocate memory for double args.");
            return;
        }
    }
    vb6DoubleArgs[vb6DoubleArgsCount - 1] = value;
    AddArgumentType(VB6_TYPE_DOUBLE);
    OutputDebugString(L"Double argument attached.");
}

DLL_EXPORT void WINAPI AttachInteger(WORD value)
{
    if (vb6IntegerArgsCount == 0)
    {
        vb6IntegerArgs = malloc(sizeof(WORD));
        if (vb6IntegerArgs == NULL)
        {
            OutputDebugString(L"Could not allocate memory for integer args.");
            return;
        }
        vb6IntegerArgsCount++;
    }
    else
    {
        WORD* originalIntegerArgs = vb6IntegerArgs;
        vb6IntegerArgs = realloc(vb6IntegerArgs, sizeof(WORD) * ++vb6IntegerArgsCount);
        if (vb6IntegerArgs == NULL)
        {
            vb6IntegerArgs = originalIntegerArgs;
            OutputDebugString(L"Could not reallocate memory for integer args.");
            return;
        }
    }
    vb6IntegerArgs[vb6IntegerArgsCount - 1] = value;
    AddArgumentType(VB6_TYPE_INTEGER);
    OutputDebugString(L"Integer argument attached.");
}

DLL_EXPORT void WINAPI AttachSingle(float value)
{
    if (vb6SingleArgsCount == 0)
    {
        vb6SingleArgs = malloc(sizeof(float));
        if (vb6SingleArgs == NULL)
        {
            OutputDebugString(L"Could not allocate memory for single args.");
            return;
        }
        vb6SingleArgsCount++;
    }
    else
    {
        float* originalSingleArgs = vb6SingleArgs;
        vb6SingleArgs = realloc(vb6SingleArgs, sizeof(float) * ++vb6SingleArgsCount);
        if (vb6SingleArgs == NULL)
        {
            vb6SingleArgs = originalSingleArgs;
            OutputDebugString(L"Could not reallocate memory for single args.");
            return;
        }
    }
    vb6SingleArgs[vb6SingleArgsCount - 1] = value;
    AddArgumentType(VB6_TYPE_SINGLE);
    OutputDebugString(L"Single argument attached.");
}

DLL_EXPORT void WINAPI AttachString(BSTR value)
{
    UINT length;
    wchar_t *pwcString;
    unsigned int i;

    length = SysStringByteLen(value);
    /* Additional space for the NULL character */
    pwcString = malloc(sizeof(wchar_t) * (length + 1));
    if (pwcString == NULL)
    {
        OutputDebugString(L"Memory allocation failed for wide string.");
        return;
    }

    /* ANSI to wide characters */
    for (i = 0; i < length; i++)
    {
        pwcString[i] = (wchar_t)(((LPSTR)value)[i]);
    }
    pwcString[length] = L'\0';

    if (vb6StringArgsCount == 0)
    {
        vb6StringArgs = malloc(sizeof(BSTR));
        if (vb6StringArgs == NULL)
        {
            OutputDebugString(L"Could not allocate memory for string args.");
            return;
        }
        vb6StringArgsCount++;
    }
    else
    {
        BSTR *originalStringArgs = vb6StringArgs;
        vb6StringArgs = realloc(vb6StringArgs, sizeof(BSTR) * ++vb6StringArgsCount);
        if (vb6StringArgs == NULL)
        {
            vb6StringArgs = originalStringArgs;
            OutputDebugString(L"Could not reallocate memory for string args.");
            return;
        }
    }
    vb6StringArgs[vb6StringArgsCount - 1] = SysAllocString(pwcString);
    AddArgumentType(VB6_TYPE_STRING);
    OutputDebugString(L"String argument attached.");

    SysFreeString(value);  /* This was allocated by VB */
    free(pwcString);
}

DLL_EXPORT void WINAPI CallByAddress(DWORD dwAddress)
{
    int i;
    double dValue;
    DWORD dwValue;
    FLOAT fValue;
    BSTR bstrValue;

    for (i = GetAttachedArgsCount(); i >= 0; i--)
    {
        switch (attachedArgsTypes[i])
        {
        case VB6_TYPE_BYTE:
        {
            /* Even 8-bit values are passed as 32-bit values. */
            dwValue = (DWORD)vb6ByteArgs[--vb6ByteArgsCount];
            __asm { PUSH dwValue };
            break;
        }

        case VB6_TYPE_DOUBLE:
        {
            dValue = vb6DoubleArgs[--vb6DoubleArgsCount];
            __asm {
                FLD dValue
                SUB ESP, 8
                FSTP QWORD PTR[ESP]
            };
            break;
        }

        case VB6_TYPE_INTEGER:
        {
            dwValue = (DWORD)vb6IntegerArgs[--vb6IntegerArgsCount];
            __asm { PUSH dwValue };
            break;
        }

        case VB6_TYPE_SINGLE:
        {
            fValue = vb6SingleArgs[--vb6SingleArgsCount];
            __asm { PUSH fValue };
            break;
        }

        case VB6_TYPE_STRING:
        {
            bstrValue = vb6StringArgs[--vb6StringArgsCount];
            __asm { PUSH bstrValue };
            break;
        }
        }
    }
    __asm {
        CALL dwAddress
    }
    /* Release allocated memory */
    free(vb6ByteArgs);
    free(vb6DoubleArgs);
    free(vb6IntegerArgs);
    free(vb6SingleArgs);
    free(vb6StringArgs);
    ResetArgs();
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
        case DLL_PROCESS_ATTACH:
            break;

        case DLL_PROCESS_DETACH:
            break;

        case DLL_THREAD_ATTACH:
            break;

        case DLL_THREAD_DETACH:
            break;
    }
    return TRUE;
}
