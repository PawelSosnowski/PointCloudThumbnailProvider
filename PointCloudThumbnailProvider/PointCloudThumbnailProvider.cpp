#include "happly.h"
#include <Shlwapi.h>
#include <thumbcache.h> // For IThumbnailProvider.
#include <gdiplus.h>
#include <iostream>

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
    *pdwAlpha = WTSAT_RGB;
    STATSTG statstg;
    HRESULT hr = _pStream->Stat(&statstg, STATFLAG_NONAME);
    if (SUCCEEDED(hr)) {
        auto fileSize = statstg.cbSize.QuadPart;
        std::vector<char> buffer(fileSize + 1);
        hr = _pStream->Read(buffer.data(), fileSize, nullptr);
        if (SUCCEEDED(hr)) {
            std::string ret(buffer.data());
            std::stringstream ss;
            ss << ret;

            happly::PLYData plyIn(ss);
            std::vector<std::array<double, 3>> vPos = plyIn.getVertexPositions();

            // tu trzeba stworzyc te bitmape z tych punktow w vPos

            Gdiplus::GdiplusStartupInput startupInput;
            ULONG_PTR gdiplusToken;
            Gdiplus::GdiplusStartup(&gdiplusToken, &startupInput, nullptr);
            Gdiplus::Bitmap* gdiPlusBitmap = new Gdiplus::Bitmap(256, 256, PixelFormat24bppRGB);

            gdiPlusBitmap->SetPixel(10, 10, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(11, 11, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(12, 12, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(13, 13, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(14, 14, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(15, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(16, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(17, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(18, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(19, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(20, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(21, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(22, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(23, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(24, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(25, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(26, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(27, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(28, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(29, 15, Gdiplus::Color::Red);
            gdiPlusBitmap->SetPixel(30, 15, Gdiplus::Color::Red);

            Gdiplus::Status status = gdiPlusBitmap->GetHBITMAP(Gdiplus::Color::Black, phbmp);
            delete gdiPlusBitmap;
            if (status == Gdiplus::Status::Ok)
                hr = S_OK;
            else
                hr = E_FAIL;

            Gdiplus::GdiplusShutdown(gdiplusToken); // shut down GDIPlus
            return hr;
        }
    }
    return E_FAIL;
}
