import ctypes
import ctypes.wintypes as wintypes

import comtypes

from capture.src.d3d import (
    D3D11_TEXTURE2D_DESC,
    ID3D11Device,
    ID3D11DeviceContext,
    ID3D11Texture2D
)


class DISPLAY_DEVICE(ctypes.Structure):
    _fields_ = [
        ("cb", wintypes.DWORD),
        ("DeviceName", wintypes.WCHAR * 32),
        ("DeviceString", wintypes.WCHAR * 128),
        ("StateFlags", wintypes.DWORD),
        ("DeviceID", wintypes.WCHAR * 128),
        ("DeviceKey", wintypes.WCHAR * 128),
    ]


class LUID(ctypes.Structure):
    _fields_ = [("LowPart", wintypes.DWORD), ("HighPart", wintypes.LONG)]


class DXGI_ADAPTER_DESC1(ctypes.Structure):
    _fields_ = [
        ("Description", wintypes.WCHAR * 128),
        ("VendorId", wintypes.UINT),
        ("DeviceId", wintypes.UINT),
        ("SubSysId", wintypes.UINT),
        ("Revision", wintypes.UINT),
        ("DedicatedVideoMemory", wintypes.ULARGE_INTEGER),
        ("DedicatedSystemMemory", wintypes.ULARGE_INTEGER),
        ("SharedSystemMemory", wintypes.ULARGE_INTEGER),
        ("AdapterLuid", LUID),
        ("Flags", wintypes.UINT),
    ]


class DXGI_OUTPUT_DESC(ctypes.Structure):
    _fields_ = [
        ("DeviceName", wintypes.WCHAR * 32),
        ("DesktopCoordinates", wintypes.RECT),
        ("AttachedToDesktop", wintypes.BOOL),
        ("Rotation", wintypes.UINT),
        ("Monitor", wintypes.HMONITOR),
    ]


class DXGI_OUTDUPL_POINTER_POSITION(ctypes.Structure):
    _fields_ = [("Position", wintypes.POINT), ("Visible", wintypes.BOOL)]


class DXGI_OUTDUPL_FRAME_INFO(ctypes.Structure):
    _fields_ = [
        ("LastPresentTime", wintypes.LARGE_INTEGER),
        ("LastMouseUpdateTime", wintypes.LARGE_INTEGER),
        ("AccumulatedFrames", wintypes.UINT),
        ("RectsCoalesced", wintypes.BOOL),
        ("ProtectedContentMaskedOut", wintypes.BOOL),
        ("PointerPosition", DXGI_OUTDUPL_POINTER_POSITION),
        ("TotalMetadataBufferSize", wintypes.UINT),
        ("PointerShapeBufferSize", wintypes.UINT),
    ]


class DXGI_MAPPED_RECT(ctypes.Structure):
    _fields_ = [("Pitch", wintypes.INT), ("pBits", ctypes.POINTER(wintypes.FLOAT))]


class IDXGIObject(comtypes.IUnknown):
    _iid_ = comtypes.GUID("{aec22fb8-76f3-4639-9be0-28eb43a67a2e}")
    _methods_ = [
        comtypes.STDMETHOD(comtypes.HRESULT, "SetPrivateData"),
        comtypes.STDMETHOD(comtypes.HRESULT, "SetPrivateDataInterface"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetPrivateData"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetParent"),
    ]


class IDXGIDeviceSubObject(IDXGIObject):
    _iid_ = comtypes.GUID("{3d3e0379-f9de-4d58-bb6c-18d62992f1a6}")
    _methods_ = [
        comtypes.STDMETHOD(comtypes.HRESULT, "GetDevice"),
    ]


class IDXGIResource(IDXGIDeviceSubObject):
    _iid_ = comtypes.GUID("{035f3ab4-482e-4e50-b41f-8a7f8bd8960b}")
    _methods_ = [
        comtypes.STDMETHOD(comtypes.HRESULT, "GetSharedHandle"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetUsage"),
        comtypes.STDMETHOD(comtypes.HRESULT, "SetEvictionPriority"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetEvictionPriority"),
    ]


class IDXGISurface(IDXGIDeviceSubObject):
    _iid_ = comtypes.GUID("{cafcb56c-6ac3-4889-bf47-9e23bbd260ec}")
    _methods_ = [
        comtypes.STDMETHOD(comtypes.HRESULT, "GetDesc"),
        comtypes.STDMETHOD(
            comtypes.HRESULT, "Map", [ctypes.POINTER(DXGI_MAPPED_RECT), wintypes.UINT]
        ),
        comtypes.STDMETHOD(comtypes.HRESULT, "Unmap"),
    ]


class IDXGIOutputDuplication(IDXGIObject):
    _iid_ = comtypes.GUID("{191cfac3-a341-470d-b26e-a864f428319c}")
    _methods_ = [
        comtypes.STDMETHOD(None, "GetDesc"),
        comtypes.STDMETHOD(
            comtypes.HRESULT,
            "AcquireNextFrame",
            [
                wintypes.UINT,
                ctypes.POINTER(DXGI_OUTDUPL_FRAME_INFO),
                ctypes.POINTER(ctypes.POINTER(IDXGIResource)),
            ],
        ),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetFrameDirtyRects"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetFrameMoveRects"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetFramePointerShape"),
        comtypes.STDMETHOD(comtypes.HRESULT, "MapDesktopSurface"),
        comtypes.STDMETHOD(comtypes.HRESULT, "UnMapDesktopSurface"),
        comtypes.STDMETHOD(comtypes.HRESULT, "ReleaseFrame"),
    ]


class IDXGIOutput(IDXGIObject):
    _iid_ = comtypes.GUID("{ae02eedb-c735-4690-8d52-5a8dc20213aa}")
    _methods_ = [
        comtypes.STDMETHOD(comtypes.HRESULT, "GetDesc", [ctypes.POINTER(DXGI_OUTPUT_DESC)]),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetDisplayModeList"),
        comtypes.STDMETHOD(comtypes.HRESULT, "FindClosestMatchingMode"),
        comtypes.STDMETHOD(comtypes.HRESULT, "WaitForVBlank"),
        comtypes.STDMETHOD(comtypes.HRESULT, "TakeOwnership"),
        comtypes.STDMETHOD(None, "ReleaseOwnership"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetGammaControlCapabilities"),
        comtypes.STDMETHOD(comtypes.HRESULT, "SetGammaControl"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetGammaControl"),
        comtypes.STDMETHOD(comtypes.HRESULT, "SetDisplaySurface"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetDisplaySurfaceData"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetFrameStatistics"),
    ]


class IDXGIOutput1(IDXGIOutput):
    _iid_ = comtypes.GUID("{00cddea8-939b-4b83-a340-a685226666cc}")
    _methods_ = [
        comtypes.STDMETHOD(comtypes.HRESULT, "GetDisplayModeList1"),
        comtypes.STDMETHOD(comtypes.HRESULT, "FindClosestMatchingMode1"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetDisplaySurfaceData1"),
        comtypes.STDMETHOD(
            comtypes.HRESULT,
            "DuplicateOutput",
            [
                ctypes.POINTER(ID3D11Device),
                ctypes.POINTER(ctypes.POINTER(IDXGIOutputDuplication)),
            ],
        ),
    ]


class IDXGIAdapter(IDXGIObject):
    _iid_ = comtypes.GUID("{2411e7e1-12ac-4ccf-bd14-9798e8534dc0}")
    _methods_ = [
        comtypes.STDMETHOD(
            comtypes.HRESULT,
            "EnumOutputs",
            [wintypes.UINT, ctypes.POINTER(ctypes.POINTER(IDXGIOutput))],
        ),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetDesc"),
        comtypes.STDMETHOD(comtypes.HRESULT, "CheckInterfaceSupport"),
    ]


class IDXGIAdapter1(IDXGIAdapter):
    _iid_ = comtypes.GUID("{29038f61-3839-4626-91fd-086879011a05}")
    _methods_ = [
        comtypes.STDMETHOD(comtypes.HRESULT, "GetDesc1", [ctypes.POINTER(DXGI_ADAPTER_DESC1)]),
    ]


class IDXGIFactory(IDXGIObject):
    _iid_ = comtypes.GUID("{7b7166ec-21c7-44ae-b21a-c9ae321ae369}")
    _methods_ = [
        comtypes.STDMETHOD(comtypes.HRESULT, "EnumAdapters"),
        comtypes.STDMETHOD(comtypes.HRESULT, "MakeWindowAssociation"),
        comtypes.STDMETHOD(comtypes.HRESULT, "GetWindowAssociation"),
        comtypes.STDMETHOD(comtypes.HRESULT, "CreateSwapChain"),
        comtypes.STDMETHOD(comtypes.HRESULT, "CreateSoftwareAdapter"),
    ]


class IDXGIFactory1(IDXGIFactory):
    _iid_ = comtypes.GUID("{770aae78-f26f-4dba-a829-253c83d1b387}")
    _methods_ = [
        comtypes.STDMETHOD(
            comtypes.HRESULT,
            "EnumAdapters1",
            [ctypes.c_uint, ctypes.POINTER(ctypes.POINTER(IDXGIAdapter1))],
        ),
        comtypes.STDMETHOD(wintypes.BOOL, "IsCurrent"),
    ]


def get_display_device_name_mapping():
    display_names = list()

    i = 0
    while True:
        display_device = DISPLAY_DEVICE()
        display_device.cb = ctypes.sizeof(display_device)

        if not ctypes.windll.user32.EnumDisplayDevicesW(None, i, ctypes.byref(display_device), 0):
            break

        if display_device.StateFlags > 0:
            is_primary_display_device = bool(display_device.StateFlags & 4)
            display_names.append((display_device.DeviceName, is_primary_display_device))

        i += 1

    display_device_name_mapping = dict()

    for display_name, is_primary in display_names:
        display_device = DISPLAY_DEVICE()
        display_device.cb = ctypes.sizeof(display_device)

        if ctypes.windll.user32.EnumDisplayDevicesW(
            display_name, 0, ctypes.byref(display_device), 0
        ):
            display_device_name_mapping[display_name.split("\\")[-1]] = (
                display_device.DeviceString,
                is_primary,
            )

    return display_device_name_mapping


def initialize_dxgi_factory():
    create_factory_func = ctypes.windll.dxgi.CreateDXGIFactory1

    create_factory_func.argtypes = (comtypes.GUID, ctypes.POINTER(ctypes.c_void_p))
    create_factory_func.restype = ctypes.c_int32

    handle = ctypes.c_void_p(0)

    create_factory_func(IDXGIFactory1._iid_, ctypes.byref(handle))
    idxgi_factory = ctypes.POINTER(IDXGIFactory1)(handle.value)

    return idxgi_factory


def discover_dxgi_adapters(dxgi_factory):
    dxgi_adapters = list()

    for i in range(0, 10):
        try:
            dxgi_adapter = ctypes.POINTER(IDXGIAdapter1)()
            dxgi_factory.EnumAdapters1(i, ctypes.byref(dxgi_adapter))

            dxgi_adapters.append(dxgi_adapter)
        except comtypes.COMError:
            break

    return dxgi_adapters


def discover_dxgi_outputs(dxgi_adapter):
    dxgi_outputs = list()

    for i in range(0, 10):
        try:
            dxgi_output = ctypes.POINTER(IDXGIOutput1)()
            dxgi_adapter.EnumOutputs(i, ctypes.byref(dxgi_output))

            dxgi_outputs.append(dxgi_output)
        except comtypes.COMError:
            break

    return dxgi_outputs


def describe_dxgi_output(dxgi_output):
    dxgi_output_description = DXGI_OUTPUT_DESC()
    dxgi_output.GetDesc(ctypes.byref(dxgi_output_description))

    rotation_mapping = {0: 0, 1: 0, 2: 90, 3: 180, 4: 270}

    return {
        "name": dxgi_output_description.DeviceName.split("\\")[-1],
        "position": {
            "left": dxgi_output_description.DesktopCoordinates.left,
            "top": dxgi_output_description.DesktopCoordinates.top,
            "right": dxgi_output_description.DesktopCoordinates.right,
            "bottom": dxgi_output_description.DesktopCoordinates.bottom,
        },
        "resolution": (
            (
                dxgi_output_description.DesktopCoordinates.right
                - dxgi_output_description.DesktopCoordinates.left
            ),
            (
                dxgi_output_description.DesktopCoordinates.bottom
                - dxgi_output_description.DesktopCoordinates.top
            ),
        ),
        "rotation": rotation_mapping.get(dxgi_output_description.Rotation, 0),
        "is_attached_to_desktop": bool(dxgi_output_description.AttachedToDesktop),
    }


def initialize_d3d_device(dxgi_adapter):
    initialize_func = ctypes.windll.d3d11.D3D11CreateDevice

    feature_levels = [45312, 45056, 41216, 40960, 37632, 37376, 37120]

    d3d_device = ctypes.POINTER(ID3D11Device)()
    d3d_device_context = ctypes.POINTER(ID3D11DeviceContext)()

    initialize_func(
        dxgi_adapter,
        0,
        None,
        0,
        ctypes.byref((ctypes.c_uint * 7)(*feature_levels)),
        len(feature_levels),
        7,
        ctypes.byref(d3d_device),
        None,
        ctypes.byref(d3d_device_context),
    )

    return d3d_device, d3d_device_context


def prepare_d3d11_texture_2d_for_cpu(d3d11_texture_2d, d3d_device):
    d3d11_texture_2d_description = D3D11_TEXTURE2D_DESC()
    d3d11_texture_2d.GetDesc(ctypes.byref(d3d11_texture_2d_description))

    d3d11_texture_2d_description_cpu = D3D11_TEXTURE2D_DESC()

    d3d11_texture_2d_description_cpu.Width = d3d11_texture_2d_description.Width
    d3d11_texture_2d_description_cpu.Height = d3d11_texture_2d_description.Height
    d3d11_texture_2d_description_cpu.MipLevels = 1
    d3d11_texture_2d_description_cpu.ArraySize = 1
    d3d11_texture_2d_description_cpu.SampleDesc.Count = 1
    d3d11_texture_2d_description_cpu.SampleDesc.Quality = 0
    d3d11_texture_2d_description_cpu.Usage = 3
    d3d11_texture_2d_description_cpu.Format = d3d11_texture_2d_description.Format
    d3d11_texture_2d_description_cpu.BindFlags = 0
    d3d11_texture_2d_description_cpu.CPUAccessFlags = 131072
    d3d11_texture_2d_description_cpu.MiscFlags = 0

    d3d11_texture_2d_cpu = ctypes.POINTER(ID3D11Texture2D)()
    d3d_device.CreateTexture2D(
        ctypes.byref(d3d11_texture_2d_description_cpu), None, ctypes.byref(d3d11_texture_2d_cpu),
    )

    return d3d11_texture_2d_cpu


def initialize_dxgi_output_duplication(dxgi_output, d3d_device):
    dxgi_output_duplication = ctypes.POINTER(IDXGIOutputDuplication)()
    dxgi_output.DuplicateOutput(d3d_device, ctypes.byref(dxgi_output_duplication))

    return dxgi_output_duplication


def get_dxgi_output_duplication_frame(dxgi_output_duplication, d3d_device, height=0):
    dxgi_output_duplication_frame_information = DXGI_OUTDUPL_FRAME_INFO()
    dxgi_resource = ctypes.POINTER(IDXGIResource)()

    dxgi_output_duplication.AcquireNextFrame(
        0, ctypes.byref(dxgi_output_duplication_frame_information), ctypes.byref(dxgi_resource),
    )

    frame = None

    if dxgi_output_duplication_frame_information.LastPresentTime > 0:
        id3d11_texture_2d = dxgi_resource.QueryInterface(ID3D11Texture2D)
        id3d11_texture_2d_cpu = prepare_d3d11_texture_2d_for_cpu(id3d11_texture_2d, d3d_device)

        d3d_device_context = ctypes.POINTER(ID3D11DeviceContext)()
        d3d_device.GetImmediateContext(ctypes.byref(d3d_device_context))

        d3d_device_context.CopyResource(id3d11_texture_2d_cpu, id3d11_texture_2d)

        id3d11_surface = id3d11_texture_2d_cpu.QueryInterface(IDXGISurface)
        dxgi_mapped_rect = DXGI_MAPPED_RECT()

        id3d11_surface.Map(ctypes.byref(dxgi_mapped_rect), 1)

        pointer = dxgi_mapped_rect.pBits
        pitch = int(dxgi_mapped_rect.Pitch)

        size = pitch * height

        frame = ctypes.string_at(pointer, size=size)
        id3d11_surface.Unmap()

    dxgi_output_duplication.ReleaseFrame()

    return frame
