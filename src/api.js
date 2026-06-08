/**
 * 发票酱 — 统一 API 适配层
 * 桌面版: 通过嵌入式 Axum 服务器 fetch() 调用
 * Web 版: 直接 fetch() 调用远程服务器
 * Tauri IPC 仅用于窗口操作等不可替代的功能
 */
var __api = (function () {
    var _serverPort = window.__serverPort || 3000;
    var _baseUrl = window.__apiBaseUrl || ('http://127.0.0.1:' + _serverPort);
    var _sessionId = '';
    var _isDesktop = !!window.__TAURI_INTERNALS__;

    // Commands that MUST use Tauri IPC (no HTTP equivalent)
    var _ipcOnly = [
        'show_window',
        'plugin:dialog|open',
        'plugin:dialog|save',
        'plugin:event|listen',
        'plugin:event|unlisten',
        'download_pdfium_dll',
        'open_invoice_files',  // Desktop: file dialog → direct IPC (no multipart HTTP)
    ];

    function _isIpcCommand(cmd) {
        return _ipcOnly.indexOf(cmd) >= 0;
    }

    function _mapCommandToEndpoint(cmd) {
        // Map Tauri command names to REST endpoints
        var map = {
            'render_pdf_pages': '/api/v1/render_pdf',
            'render_pdf_pages_pdfium': '/api/v1/render_pdf',
            'extract_pdf_text': '/api/v1/extract_pdf_text',
            'extract_pdf_texts': '/api/v1/extract_pdf_texts',
            'generate_pdf_from_layout': '/api/v1/generate_pdf',
            'get_printers': '/api/v1/printers',
            'list_printers': '/api/v1/printers',
            'pdfium_print': '/api/v1/print',
            'pdfium_print_pdf': '/api/v1/print',
            'pdfium_vector_print': '/api/v1/print',
            'print_pdf_file': '/api/v1/print',
            'shell_execute_print': '/api/v1/print',
            'ocr_image': '/api/v1/ocr_image',
            'ocr_pdf_page': '/api/v1/ocr_pdf_page',
            'parse_ofd': '/api/v1/parse_ofd',
            'parse_xml_invoice': '/api/v1/parse_xml_invoice',
            'open_ofd_images': '/api/v1/open_ofd_images',
            'open_invoice_files': null,  // Desktop-only: file dialog → IPC invoke
            'check_path_exists': '/api/v1/check_path_exists',
            'check_ocr_available': null,  // Desktop-only: OCR availability check
            'check_winrt_pdf': null,  // Desktop-only: WinRT PDF check
            'check_pdfium_available': '/api/v1/health',
            'get_config': '/api/v1/get_config',
            'get_app_version': '/api/v1/get_app_version',
            'cancel_download': '/api/v1/cancel_download',
            'trim_image': '/api/v1/trim_image',
            'copy_file': '/api/v1/copy_file',
            'rename_file': '/api/v1/rename_file',
            'write_text_file': '/api/v1/write_text_file',
            'get_temp_dir': '/api/v1/get_temp_dir',
            'get_downloads_dir': '/api/v1/get_downloads_dir',
            'download_pdfium_dll': '/api/v1/health', // Web: no download needed, health check confirms availability
            'open_file': null,   // Desktop-only: open file with system app
            'open_url': null,    // Desktop-only: open URL in browser
        };
        return map[cmd] || null;
    }

    function _addSessionId(params) {
        // Always include sessionId field to satisfy server-side struct deserialization.
        // Desktop mode uses empty string (default "_desktop" session), Web mode uses real session id.
        params.sessionId = _sessionId || '';
        return params;
    }

    async function call(cmd, params) {
        // IPC-only commands: use Tauri invoke directly
        if (_isIpcCommand(cmd) && _isDesktop) {
            return _invokeTauri(cmd, params);
        }

        // Special handling: check_ocr_available in Web mode returns false (no OCR in web build)
        if (cmd === 'check_ocr_available') {
            if (_isDesktop) return _invokeTauri(cmd, params);
            return false;
        }

        var endpoint = _mapCommandToEndpoint(cmd);

        // Desktop-only commands (mapped to null): fallback to Tauri IPC on desktop, error on web
        if (endpoint === null) {
            if (_isDesktop) {
                return _invokeTauri(cmd, params);
            }
            throw new Error('Command not available in Web mode: ' + cmd);
        }

        // Special handling for health check (returns different format)
        if (cmd === 'check_pdfium_available') {
            var resp = await _fetch(endpoint, {});
            return resp.data && resp.data.pdfium === true;
        }

        // Special handling: download_pdfium_dll in Web mode is a no-op
        if (cmd === 'download_pdfium_dll' && !_isDesktop) {
            return { success: true, message: 'Web 版本无需下载 PDFium' };
        }

        // Standard POST request
        var body = _addSessionId(Object.assign({}, params || {}));
        var result = await _fetch(endpoint, body);

        if (!result.ok) {
            var err = new Error(result.error || 'Unknown error');
            err.code = result.code;
            err.recoverable = result.recoverable;
            throw err;
        }

        return result.data;
    }

    async function _fetch(endpoint, body) {
        var url = _baseUrl + endpoint;
        var opts = {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        };

        // For GET endpoints
        if (endpoint === '/api/v1/printers' || endpoint === '/api/v1/health') {
            opts.method = 'GET';
            delete opts.body;
        }

        var response = await fetch(url, opts);
        if (!response.ok) {
            var errorBody;
            try {
                errorBody = await response.json();
            } catch (_) {
                throw new Error('HTTP ' + response.status + ': ' + response.statusText);
            }
            if (errorBody && errorBody.error) {
                var err = new Error(errorBody.error);
                err.code = errorBody.code;
                err.recoverable = errorBody.recoverable;
                throw err;
            }
            throw new Error('HTTP ' + response.status);
        }

        return await response.json();
    }

    async function _uploadFiles(params) {
        // Desktop: open_invoice_files is handled by IPC (see _ipcOnly);
        // this function is only called for Web multipart uploads.
        throw new Error('Use uploadFilesFromInput() for Web mode');
    }

    async function uploadFilesFromInput(fileList) {
        var formData = new FormData();
        for (var i = 0; i < fileList.length; i++) {
            formData.append('files', fileList[i]);
        }

        var url = _baseUrl + '/api/v1/upload';
        var response = await fetch(url, {
            method: 'POST',
            body: formData,
        });

        if (!response.ok) {
            throw new Error('Upload failed: HTTP ' + response.status);
        }

        var result = await response.json();
        if (result.ok && result.data) {
            _sessionId = result.data.sessionId || _sessionId;
            return result.data.files;
        }
        throw new Error(result.error || 'Upload failed');
    }

    async function _invokeTauri(cmd, params) {
        if (typeof window.__TAURI_INVOKE__ === 'function') {
            return window.__TAURI_INVOKE__(cmd, params);
        }
        // Tauri 2.x style
        if (window.__TAURI_INTERNALS__) {
            return window.__TAURI_INTERNALS__.invoke(cmd, params);
        }
        throw new Error('Tauri IPC not available');
    }

    function setServerPort(port) {
        _serverPort = port;
        _baseUrl = 'http://127.0.0.1:' + port;
    }

    function setSessionId(id) {
        _sessionId = id;
    }

    function getSessionId() {
        return _sessionId;
    }

    function isDesktop() {
        return _isDesktop;
    }

    function getBaseUrl() {
        return _baseUrl;
    }

    // Wait for embedded server to be ready (desktop mode: poll /api/v1/health up to 5s)
    var _serverReady = false;
    async function init() {
        if (!_isDesktop) { _serverReady = true; return true; }
        for (var i = 0; i < 50; i++) {
            try {
                var resp = await fetch('http://127.0.0.1:' + _serverPort + '/api/v1/health');
                if (resp.ok) { _serverReady = true; return true; }
            } catch(e) {}
            await new Promise(function(r) { setTimeout(r, 100); });
        }
        console.warn('Embedded server not ready after 5s');
        return false;
    }

    // SSE progress listener for long-running tasks (e.g., PDF generation)
    function listen(taskId, callback) {
        var url = (_isDesktop ? 'http://127.0.0.1:' + _serverPort : '') + '/api/v1/progress/' + taskId;
        var source = new EventSource(url);
        source.onmessage = function(e) {
            try { callback(JSON.parse(e.data)); } catch(ex) { callback(e.data); }
        };
        source.onerror = function() { source.close(); };
        return { close: function() { source.close(); } };
    }

    // Download file: desktop opens with system app, web triggers browser download
    function downloadFile(sessionId, filename) {
        if (_isDesktop) {
            call('open_file', { path: filename });
        } else {
            var a = document.createElement('a');
            a.href = _baseUrl + '/api/v1/download/' + sessionId + '/' + encodeURIComponent(filename);
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }
    }

    // Open URL: desktop uses system browser, web uses window.open
    function openUrl(url) {
        if (_isDesktop) {
            call('open_url', { url: url });
        } else {
            window.open(url, '_blank');
        }
    }

    return {
        init: init,
        call: call,
        listen: listen,
        uploadFilesFromInput: uploadFilesFromInput,
        downloadFile: downloadFile,
        openUrl: openUrl,
        setServerPort: setServerPort,
        setSessionId: setSessionId,
        getSessionId: getSessionId,
        isDesktop: isDesktop,
        getBaseUrl: getBaseUrl,
    };
})();
