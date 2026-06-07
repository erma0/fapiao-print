#include <windows.h>

typedef void (*FnRenderPage)(void* dc, void* page, int x, int y, int w, int h, int rotate, int flags);

DWORD SafeCallRenderPage(FnRenderPage func, void* dc, void* page, int x, int y, int w, int h, int rotate, int flags) {
    __try {
        func(dc, page, x, y, w, h, rotate, flags);
        return 0;
    } __except(EXCEPTION_EXECUTE_HANDLER) {
        return GetExceptionCode();
    }
}
