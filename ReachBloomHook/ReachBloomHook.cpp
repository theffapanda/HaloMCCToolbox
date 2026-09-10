#include <windows.h>
#include <d3d11.h>
#include <atomic>
#include <mutex>
#include <unordered_map>
#include "MinHook.h"

#pragma comment(lib, "d3d11.lib")

namespace
{
    constexpr uint32_t ReachBloomShaderCrc = 981019040U;
    constexpr uint32_t StateMagic = 0x52424C4D; // RBLM

    struct SharedState
    {
        uint32_t magic;
        uint32_t version;
        volatile LONG installOk;
        volatile LONG shaderFound;
        volatile LONG bloomDisabled;
        volatile LONG64 blockedDraws;
    };

    using CreatePixelShaderFn = HRESULT(STDMETHODCALLTYPE *)(ID3D11Device *, const void *, SIZE_T, ID3D11ClassLinkage *, ID3D11PixelShader **);
    using PSSetShaderFn = void(STDMETHODCALLTYPE *)(ID3D11DeviceContext *, ID3D11PixelShader *, ID3D11ClassInstance *const *, UINT);
    using DrawFn = void(STDMETHODCALLTYPE *)(ID3D11DeviceContext *, UINT, UINT);
    using DrawIndexedFn = void(STDMETHODCALLTYPE *)(ID3D11DeviceContext *, UINT, UINT, INT);

    CreatePixelShaderFn originalCreatePixelShader = nullptr;
    PSSetShaderFn originalSetImmediate = nullptr;
    PSSetShaderFn originalSetDeferred = nullptr;
    DrawFn originalDrawImmediate = nullptr;
    DrawFn originalDrawDeferred = nullptr;
    DrawIndexedFn originalDrawIndexedImmediate = nullptr;
    DrawIndexedFn originalDrawIndexedDeferred = nullptr;

    std::mutex stateMutex;
    std::unordered_map<ID3D11PixelShader *, uint32_t> shaderHashes;
    std::unordered_map<ID3D11DeviceContext *, uint32_t> activePixelShaders;
    std::atomic<bool> disableBloom{false};

    HANDLE stateMapping = nullptr;
    HANDLE disableEvent = nullptr;
    HANDLE enableEvent = nullptr;
    SharedState *sharedState = nullptr;

    uint32_t Crc32(const uint8_t *data, size_t size)
    {
        uint32_t crc = 0xFFFFFFFFU;
        for (size_t i = 0; i < size; ++i)
        {
            crc ^= data[i];
            for (int bit = 0; bit < 8; ++bit)
                crc = (crc >> 1) ^ (0xEDB88320U & (0U - (crc & 1U)));
        }
        return ~crc;
    }

    HRESULT STDMETHODCALLTYPE HookCreatePixelShader(
        ID3D11Device *device, const void *bytecode, SIZE_T length,
        ID3D11ClassLinkage *linkage, ID3D11PixelShader **shader)
    {
        const HRESULT result = originalCreatePixelShader(device, bytecode, length, linkage, shader);
        if (SUCCEEDED(result) && shader && *shader && bytecode && length)
        {
            const uint32_t crc = Crc32(static_cast<const uint8_t *>(bytecode), length);
            {
                std::lock_guard lock(stateMutex);
                shaderHashes[*shader] = crc;
            }
            if (crc == ReachBloomShaderCrc && sharedState)
                InterlockedExchange(&sharedState->shaderFound, 1);
        }
        return result;
    }

    void ProcessSetShader(
        PSSetShaderFn original, ID3D11DeviceContext *context,
        ID3D11PixelShader *shader, ID3D11ClassInstance *const *classes, UINT count)
    {
        uint32_t crc = 0;
        {
            std::lock_guard lock(stateMutex);
            const auto found = shaderHashes.find(shader);
            if (found != shaderHashes.end())
                crc = found->second;
            activePixelShaders[context] = crc;
        }
        original(context, shader, classes, count);
    }

    bool ShouldBlock(ID3D11DeviceContext *context)
    {
        if (!disableBloom.load(std::memory_order_relaxed))
            return false;

        std::lock_guard lock(stateMutex);
        const auto found = activePixelShaders.find(context);
        return found != activePixelShaders.end() && found->second == ReachBloomShaderCrc;
    }

    void RecordBlockedDraw()
    {
        if (sharedState)
            InterlockedIncrement64(&sharedState->blockedDraws);
    }

    void STDMETHODCALLTYPE HookSetImmediate(ID3D11DeviceContext *context, ID3D11PixelShader *shader, ID3D11ClassInstance *const *classes, UINT count)
    {
        ProcessSetShader(originalSetImmediate, context, shader, classes, count);
    }

    void STDMETHODCALLTYPE HookSetDeferred(ID3D11DeviceContext *context, ID3D11PixelShader *shader, ID3D11ClassInstance *const *classes, UINT count)
    {
        ProcessSetShader(originalSetDeferred, context, shader, classes, count);
    }

    void STDMETHODCALLTYPE HookDrawImmediate(ID3D11DeviceContext *context, UINT vertexCount, UINT startVertex)
    {
        if (ShouldBlock(context)) { RecordBlockedDraw(); return; }
        originalDrawImmediate(context, vertexCount, startVertex);
    }

    void STDMETHODCALLTYPE HookDrawDeferred(ID3D11DeviceContext *context, UINT vertexCount, UINT startVertex)
    {
        if (ShouldBlock(context)) { RecordBlockedDraw(); return; }
        originalDrawDeferred(context, vertexCount, startVertex);
    }

    void STDMETHODCALLTYPE HookDrawIndexedImmediate(ID3D11DeviceContext *context, UINT indexCount, UINT startIndex, INT baseVertex)
    {
        if (ShouldBlock(context)) { RecordBlockedDraw(); return; }
        originalDrawIndexedImmediate(context, indexCount, startIndex, baseVertex);
    }

    void STDMETHODCALLTYPE HookDrawIndexedDeferred(ID3D11DeviceContext *context, UINT indexCount, UINT startIndex, INT baseVertex)
    {
        if (ShouldBlock(context)) { RecordBlockedDraw(); return; }
        originalDrawIndexedDeferred(context, indexCount, startIndex, baseVertex);
    }

    bool AddHook(void *target, void *detour, void **original)
    {
        return target && MH_CreateHook(target, detour, original) == MH_OK && MH_EnableHook(target) == MH_OK;
    }

    bool CreateControlChannel()
    {
        wchar_t name[128]{};
        const DWORD pid = GetCurrentProcessId();
        swprintf_s(name, L"Local\\HaloMCCToolbox.ReachBloom.State.%lu", pid);
        stateMapping = CreateFileMappingW(INVALID_HANDLE_VALUE, nullptr, PAGE_READWRITE, 0, sizeof(SharedState), name);
        if (!stateMapping)
            return false;

        sharedState = static_cast<SharedState *>(MapViewOfFile(stateMapping, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(SharedState)));
        if (!sharedState)
            return false;

        ZeroMemory(sharedState, sizeof(SharedState));
        sharedState->magic = StateMagic;
        sharedState->version = 1;

        swprintf_s(name, L"Local\\HaloMCCToolbox.ReachBloom.Disable.%lu", pid);
        disableEvent = CreateEventW(nullptr, FALSE, FALSE, name);
        swprintf_s(name, L"Local\\HaloMCCToolbox.ReachBloom.Enable.%lu", pid);
        enableEvent = CreateEventW(nullptr, FALSE, FALSE, name);
        return disableEvent && enableEvent;
    }

    DWORD WINAPI Install(void *)
    {
        if (!CreateControlChannel())
            return 1;

        ID3D11Device *device = nullptr;
        ID3D11DeviceContext *immediate = nullptr;
        ID3D11DeviceContext *deferred = nullptr;
        D3D_FEATURE_LEVEL obtained{};
        D3D_FEATURE_LEVEL requested[] = {D3D_FEATURE_LEVEL_11_0};
        HRESULT result = D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, 0, requested, 1, D3D11_SDK_VERSION, &device, &obtained, &immediate);
        if (FAILED(result))
            result = D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_WARP, nullptr, 0, requested, 1, D3D11_SDK_VERSION, &device, &obtained, &immediate);
        if (FAILED(result) || !device || !immediate || MH_Initialize() != MH_OK)
            return 2;

        device->CreateDeferredContext(0, &deferred);
        void **deviceVtable = *reinterpret_cast<void ***>(device);
        void **immediateVtable = *reinterpret_cast<void ***>(immediate);
        void **deferredVtable = deferred ? *reinterpret_cast<void ***>(deferred) : nullptr;

        bool installed = AddHook(deviceVtable[15], reinterpret_cast<void *>(&HookCreatePixelShader), reinterpret_cast<void **>(&originalCreatePixelShader));
        installed &= AddHook(immediateVtable[9], reinterpret_cast<void *>(&HookSetImmediate), reinterpret_cast<void **>(&originalSetImmediate));
        installed &= AddHook(immediateVtable[13], reinterpret_cast<void *>(&HookDrawImmediate), reinterpret_cast<void **>(&originalDrawImmediate));
        installed &= AddHook(immediateVtable[12], reinterpret_cast<void *>(&HookDrawIndexedImmediate), reinterpret_cast<void **>(&originalDrawIndexedImmediate));

        if (deferredVtable && deferredVtable[9] != immediateVtable[9])
            installed &= AddHook(deferredVtable[9], reinterpret_cast<void *>(&HookSetDeferred), reinterpret_cast<void **>(&originalSetDeferred));
        else originalSetDeferred = originalSetImmediate;
        if (deferredVtable && deferredVtable[13] != immediateVtable[13])
            installed &= AddHook(deferredVtable[13], reinterpret_cast<void *>(&HookDrawDeferred), reinterpret_cast<void **>(&originalDrawDeferred));
        else originalDrawDeferred = originalDrawImmediate;
        if (deferredVtable && deferredVtable[12] != immediateVtable[12])
            installed &= AddHook(deferredVtable[12], reinterpret_cast<void *>(&HookDrawIndexedDeferred), reinterpret_cast<void **>(&originalDrawIndexedDeferred));
        else originalDrawIndexedDeferred = originalDrawIndexedImmediate;

        if (deferred) deferred->Release();
        immediate->Release();
        device->Release();

        InterlockedExchange(&sharedState->installOk, installed ? 1 : -1);
        if (!installed)
            return 3;

        HANDLE controls[] = {disableEvent, enableEvent};
        for (;;)
        {
            const DWORD wait = WaitForMultipleObjects(2, controls, FALSE, INFINITE);
            if (wait == WAIT_OBJECT_0)
            {
                disableBloom.store(true, std::memory_order_relaxed);
                InterlockedExchange(&sharedState->bloomDisabled, 1);
            }
            else if (wait == WAIT_OBJECT_0 + 1)
            {
                disableBloom.store(false, std::memory_order_relaxed);
                InterlockedExchange(&sharedState->bloomDisabled, 0);
            }
        }
    }
}

BOOL APIENTRY DllMain(HMODULE module, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        DisableThreadLibraryCalls(module);
        const HANDLE thread = CreateThread(nullptr, 0, Install, nullptr, 0, nullptr);
        if (thread) CloseHandle(thread);
    }
    return TRUE;
}

