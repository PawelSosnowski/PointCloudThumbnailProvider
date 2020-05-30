#include <Shlwapi.h>
#include <thumbcache.h> // For IThumbnailProvider.
#include <gdiplus.h>
#include <new>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "gdiplus.lib")


// this thumbnail provider implements IInitializeWithStream to enable being hosted
// in an isolated process for robustness

class CPointCloudThumbProvider : public IInitializeWithStream, IThumbnailProvider {
public:
    CPointCloudThumbProvider() : _cRef(1), _pStream(nullptr) {}

    virtual ~CPointCloudThumbProvider() {
        if (_pStream) {
            _pStream->Release();
        }
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        static const QITAB qit[] = {
                QITABENT(CPointCloudThumbProvider, IInitializeWithStream),
                QITABENT(CPointCloudThumbProvider, IThumbnailProvider),
                {nullptr}
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef() override {
        return InterlockedIncrement(&_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG cRef = InterlockedDecrement(&_cRef);
        if (!cRef) {
            delete this;
        }
        return cRef;
    }

    // IInitializeWithStream
    IFACEMETHODIMP Initialize(IStream* pStream, DWORD grfMode) override;

    // IThumbnailProvider
    IFACEMETHODIMP GetThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha) override;

private:
    HRESULT _FindThumbnailDataFromStream(ULONG* actualResourceDataSize);

    HRESULT _GetThumbnailFromPointCloudThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha,
        ULONG actualResourceDataSize);

    long _cRef;
    IStream* _pStream;     // provided during initialization.
};

HRESULT CPointCloudThumbProvider_CreateInstance(REFIID riid, void** ppv) {
    auto* pNew = new(std::nothrow) CPointCloudThumbProvider();
    HRESULT hr = pNew ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr)) {
        hr = pNew->QueryInterface(riid, ppv);
        pNew->Release();
    }
    return hr;
}

// IInitializeWithStream
IFACEMETHODIMP CPointCloudThumbProvider::Initialize(IStream* pStream, DWORD) {
    HRESULT hr = E_UNEXPECTED;  // can only be inited once
    if (_pStream == nullptr) {
        // take a reference to the stream if we have not been inited yet
        hr = pStream->QueryInterface(&_pStream);
    }
    return hr;
}

// IThumbnailProvider
IFACEMETHODIMP CPointCloudThumbProvider::GetThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha) {
    ULONG actualResourceDataSize;
    HRESULT hr = _FindThumbnailDataFromStream(&actualResourceDataSize);
    if (SUCCEEDED(hr)) {
        Gdiplus::GdiplusStartupInput startupInput;
        ULONG_PTR gdiplusToken;
        Gdiplus::GdiplusStartup(&gdiplusToken, &startupInput, nullptr);

        HRESULT hr_thumb = _GetThumbnailFromPointCloudThumbnail(cx, phbmp, pdwAlpha, actualResourceDataSize);

        Gdiplus::GdiplusShutdown(gdiplusToken); // shut down GDIPlus

        return hr_thumb;
    }
    return E_FAIL;
}

ULONG ReadUInt32BE(IStream* pStream) {
    unsigned char buf[4];
    pStream->Read(buf, 4, nullptr);

    ULONG ret = ((buf[0] << 24) | (buf[1] << 16) | (buf[2] << 8) | (buf[3] << 0));
    return ret;
}

ULONGLONG GetStreamPosition(IStream* pStream) {
    LARGE_INTEGER seek_offset_zero = { 0 };
    ULARGE_INTEGER currentPosition = { 0 };
    pStream->Seek(seek_offset_zero, STREAM_SEEK_CUR, &currentPosition);
    return currentPosition.QuadPart;
}

HRESULT CPointCloudThumbProvider::_FindThumbnailDataFromStream(ULONG* pActualResourceDataSize) {
    LARGE_INTEGER seek_offset_zero = { 0 };
    _pStream->Seek(seek_offset_zero, STREAM_SEEK_SET, nullptr);
    {
        // https://www.adobe.com/devnet-apps/photoshop/fileformatashtml/

        // file header
        LARGE_INTEGER seek_offset_file_header = { 4 + 2 + 6 + 2 + 4 + 4 + 2 + 2 };
        _pStream->Seek(seek_offset_file_header, STREAM_SEEK_CUR, nullptr);

        ULONG colorModeDataSectionLength = ReadUInt32BE(_pStream);
        LARGE_INTEGER seek_offset_colorModeDataSectionLength = { colorModeDataSectionLength };
        _pStream->Seek(seek_offset_colorModeDataSectionLength, STREAM_SEEK_CUR, nullptr);

        ULONG imageResourcesSectionLength = ReadUInt32BE(_pStream);
        ULONGLONG imageResourceFinishPos = GetStreamPosition(_pStream) + imageResourcesSectionLength;

        while (GetStreamPosition(_pStream) != imageResourceFinishPos) {
            if (GetStreamPosition(_pStream) > imageResourceFinishPos)
                return E_FAIL;

            // Signature: '8BIM'
            unsigned char buf_signature[4];
            _pStream->Read(buf_signature, 4, nullptr);

            // Unique identifier for the resource.
            unsigned char buf_imageResourceId[2];
            _pStream->Read(buf_imageResourceId, 2, nullptr);

            // 0x0409 (Photoshop 4.0) Thumbnail resource for Photoshop 4.0 only.See See Thumbnail resource format.
            // 0x040C (Photoshop 5.0) Thumbnail resource (supersedes resource 1033). See See Thumbnail resource format.
            bool isThumbnailResource = false;
            if ((buf_imageResourceId[0] == 0x04 && buf_imageResourceId[1] == 0x09) ||
                (buf_imageResourceId[0] == 0x04 && buf_imageResourceId[1] == 0x0C))
                isThumbnailResource = true;

            // Name: Pascal string, padded to make the size even (a null name consists of two bytes of 0)
            unsigned char pascalStringLength;
            _pStream->Read(&pascalStringLength, 1, nullptr);
            LARGE_INTEGER seek_offset_pascalString = {
                    static_cast<DWORD>(pascalStringLength + (pascalStringLength % 2 == 0 ? 1 : 0)) };
            _pStream->Seek(seek_offset_pascalString, STREAM_SEEK_CUR, nullptr);

            ULONG actualResourceDataSize = ReadUInt32BE(_pStream);  // Actual size of resource data that follows
            if (!isThumbnailResource) {
                // The resource data, described in the sections on the individual resource types.
                // ** It is padded to make the size even. **
                LARGE_INTEGER seek_offset_actualResource = { actualResourceDataSize + (actualResourceDataSize % 2) };
                _pStream->Seek(seek_offset_actualResource, STREAM_SEEK_CUR, nullptr);
            }
            else {
                *pActualResourceDataSize = actualResourceDataSize;
                return S_OK;
            }
        }
    }
    return E_FAIL;
}

HRESULT CPointCloudThumbProvider::_GetThumbnailFromPointCloudThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha,
    ULONG actualResourceDataSize) {
    auto* buf_actualResourceData = new unsigned char[actualResourceDataSize];
    _pStream->Read(buf_actualResourceData, actualResourceDataSize, nullptr);
    IStream* memstream = SHCreateMemStream(buf_actualResourceData + 28, actualResourceDataSize - 28);
    Gdiplus::Bitmap* gdiPlusBitmap = Gdiplus::Bitmap::FromStream(memstream);
    memstream->Release();
    delete[] buf_actualResourceData;

    Gdiplus::Status status = gdiPlusBitmap->GetHBITMAP(Gdiplus::Color::Black, phbmp);
    delete gdiPlusBitmap;
    if (status == Gdiplus::Status::Ok) {
        return S_OK;
    }
    return E_FAIL;
}
