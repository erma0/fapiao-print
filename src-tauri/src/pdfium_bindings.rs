#![allow(non_camel_case_types)]

use std::ffi::c_void;
use std::sync::LazyLock;

type FPDF_DOCUMENT = *mut c_void;
type FPDF_PAGE = *mut c_void;
type FPDF_BITMAP = *mut c_void;

type FnInitLibrary = unsafe fn();
type FnDestroyLibrary = unsafe fn();
type FnLoadMemDocument = unsafe fn(*const c_void, i32, *const u8) -> FPDF_DOCUMENT;
type FnGetPageCount = unsafe fn(FPDF_DOCUMENT) -> i32;
type FnGetPageWidthF = unsafe fn(FPDF_PAGE) -> f32;
type FnGetPageHeightF = unsafe fn(FPDF_PAGE) -> f32;
type FnLoadPage = unsafe fn(FPDF_DOCUMENT, i32) -> FPDF_PAGE;
type FnRenderPage = unsafe extern "C" fn(*mut c_void, FPDF_PAGE, i32, i32, i32, i32, i32, i32);
type FnClosePage = unsafe fn(FPDF_PAGE);
type FnCloseDocument = unsafe fn(FPDF_DOCUMENT);
type FnGetLastError = unsafe fn() -> i32;
type FnBitmapCreate = unsafe fn(i32, i32, i32) -> FPDF_BITMAP;
type FnBitmapFillRect = unsafe fn(FPDF_BITMAP, i32, i32, i32, i32, u32);
type FnRenderPageBitmap = unsafe fn(FPDF_BITMAP, FPDF_PAGE, i32, i32, i32, i32, i32, i32);
type FnBitmapGetBuffer = unsafe fn(FPDF_BITMAP) -> *mut c_void;
type FnBitmapGetStride = unsafe fn(FPDF_BITMAP) -> i32;
type FnBitmapDestroy = unsafe fn(FPDF_BITMAP);

pub const FPDF_ANNOT: i32 = 0x01;
pub const FPDF_PRINTING: i32 = 0x800;

pub const FPDF_ERR_SUCCESS: i32 = 0;
pub const FPDF_ERR_UNKNOWN: i32 = 1;
pub const FPDF_ERR_FILE: i32 = 2;
pub const FPDF_ERR_FORMAT: i32 = 3;
pub const FPDF_ERR_PASSWORD: i32 = 4;
pub const FPDF_ERR_SECURITY: i32 = 5;
pub const FPDF_ERR_PAGE: i32 = 6;

pub struct PdfiumFuncs {
    pub init_library: FnInitLibrary,
    pub _destroy_library: FnDestroyLibrary,
    pub load_mem_document: FnLoadMemDocument,
    pub get_page_count: FnGetPageCount,
    pub get_page_width_f: FnGetPageWidthF,
    pub get_page_height_f: FnGetPageHeightF,
    pub load_page: FnLoadPage,
    pub render_page: FnRenderPage,
    pub close_page: FnClosePage,
    pub close_document: FnCloseDocument,
    pub get_last_error: FnGetLastError,
    pub bitmap_create: FnBitmapCreate,
    pub bitmap_fill_rect: FnBitmapFillRect,
    pub render_page_bitmap: FnRenderPageBitmap,
    pub bitmap_get_buffer: FnBitmapGetBuffer,
    pub bitmap_get_stride: FnBitmapGetStride,
    pub bitmap_destroy: FnBitmapDestroy,
}

pub struct PdfiumState {
    pub _lib: libloading::Library,
    pub funcs: PdfiumFuncs,
}

static PDFIUM: LazyLock<std::sync::RwLock<Option<PdfiumState>>> =
    LazyLock::new(|| std::sync::RwLock::new(None));

pub fn with_pdfium<F, R>(f: F) -> Result<R, String>
where
    F: FnOnce(&PdfiumFuncs) -> Result<R, String>,
{
    // Fast path: read lock for already-initialized state
    {
        let guard = PDFIUM.read().unwrap();
        if let Some(ref state) = *guard {
            return f(&state.funcs);
        }
    }

    // Slow path: write lock to initialize
    let mut guard = PDFIUM.write().unwrap();
    if guard.is_none() {
        let lib = crate::platform::load_pdfium_lib()?;
        let funcs = get_pdfium_funcs(&lib)?;
        unsafe { (funcs.init_library)() };
        log::info!("PDFium library initialized");
        *guard = Some(PdfiumState { _lib: lib, funcs });
    }
    let state = guard.as_ref().unwrap();
    f(&state.funcs)
}

fn get_pdfium_funcs(lib: &libloading::Library) -> Result<PdfiumFuncs, String> {
    macro_rules! get_fn {
        ($name:expr, $ty:ty) => {
            unsafe {
                lib.get::<$ty>($name)
                    .map(|f| *f)
                    .map_err(|e| format!("无法获取函数 {}: {}", String::from_utf8_lossy($name), e))
            }
        };
    }
    Ok(PdfiumFuncs {
        init_library: get_fn!(b"FPDF_InitLibrary\0", FnInitLibrary)?,
        _destroy_library: get_fn!(b"FPDF_DestroyLibrary\0", FnDestroyLibrary)?,
        load_mem_document: get_fn!(b"FPDF_LoadMemDocument\0", FnLoadMemDocument)?,
        get_page_count: get_fn!(b"FPDF_GetPageCount\0", FnGetPageCount)?,
        get_page_width_f: get_fn!(b"FPDF_GetPageWidthF\0", FnGetPageWidthF)?,
        get_page_height_f: get_fn!(b"FPDF_GetPageHeightF\0", FnGetPageHeightF)?,
        load_page: get_fn!(b"FPDF_LoadPage\0", FnLoadPage)?,
        render_page: get_fn!(b"FPDF_RenderPage\0", FnRenderPage)
            .map_err(|e| format!("FPDF_RenderPage 不可用 (需要 Windows 版 PDFium): {}", e))?,
        close_page: get_fn!(b"FPDF_ClosePage\0", FnClosePage)?,
        close_document: get_fn!(b"FPDF_CloseDocument\0", FnCloseDocument)?,
        get_last_error: get_fn!(b"FPDF_GetLastError\0", FnGetLastError)?,
        bitmap_create: get_fn!(b"FPDFBitmap_Create\0", FnBitmapCreate)?,
        bitmap_fill_rect: get_fn!(b"FPDFBitmap_FillRect\0", FnBitmapFillRect)?,
        render_page_bitmap: get_fn!(b"FPDF_RenderPageBitmap\0", FnRenderPageBitmap)?,
        bitmap_get_buffer: get_fn!(b"FPDFBitmap_GetBuffer\0", FnBitmapGetBuffer)?,
        bitmap_get_stride: get_fn!(b"FPDFBitmap_GetStride\0", FnBitmapGetStride)?,
        bitmap_destroy: get_fn!(b"FPDFBitmap_Destroy\0", FnBitmapDestroy)?,
    })
}

pub fn pdfium_err_desc(code: i32) -> &'static str {
    match code {
        FPDF_ERR_SUCCESS => "成功",
        FPDF_ERR_UNKNOWN => "未知错误",
        FPDF_ERR_FILE => "文件不存在或无法读取",
        FPDF_ERR_FORMAT => "PDF格式错误",
        FPDF_ERR_PASSWORD => "需要密码",
        FPDF_ERR_SECURITY => "安全限制",
        FPDF_ERR_PAGE => "页面错误",
        _ => "未知错误码",
    }
}

pub fn find_pdfium_lib() -> Option<std::path::PathBuf> {
    crate::platform::find_pdfium_lib()
}
