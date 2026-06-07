/**
 * 发票酱 — 统一 API 适配层
 * 桌面版: 通过嵌入式 Axum 服务器 fetch() 调用
 * Web 版: 直接 fetch() 调用远程服务器
 * Tauri IPC 仅用于窗口操作等不可替代的功能
 */
var __api = (function () {
    var _serverPort = window.__serverPort || 3000;
    var _baseUrl = 'http://127.0.0.1:' + _serverPort;
    var _sessionId = '';
    var _isDesktop = !!window.__TAURI_INTERNALS__;

    // Commands that MUST use Tauri IPC (no HTTP equivalent)
    var _ipcOnly = [
        'show_window',
        'plugin:dialog|open',
        'plugin:dialog|save',
        'plugin:event|listen',
        'plugin:event|unlisten',
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
            'list_printers': '/api/v1/printers',
            'pdfium_print': '/api/v1/print',
            'sumatra_print': '/api/v1/print',
            'shell_execute_print': '/api/v1/print',
            'ocr_image': '/api/v1/ocr_image',
            'ocr_pdf_page': '/api/v1/ocr_pdf_page',
            'parse_ofd': '/api/v1/parse_ofd',
            'parse_xml_invoice': '/api/v1/parse_xml_invoice',
            'open_ofd_images': '/api/v1/open_ofd_images',
            'open_invoice_files': '/api/v1/upload',
            'check_path_exists': '/api/v1/check_path_exists',
            'get_config': '/api/v1/get_config',
            'get_app_version': '/api/v1/get_app_version',
            'cancel_download': '/api/v1/cancel_download',
            'trim_image': '/api/v1/trim_image',
            'copy_file': '/api/v1/copy_file',
            'rename_file': '/api/v1/rename_file',
            'write_text_file': '/api/v1/write_text_file',
            'get_temp_dir': '/api/v1/get_temp_dir',
            'get_downloads_dir': '/api/v1/get_downloads_dir',
            'check_pdfium_available': '/api/v1/health',
            'download_pdfium_dll': '/api/v1/cancel_download', // Web: no download needed
        };
        return map[cmd] || null;
    }

    function _addSessionId(params) {
        if (_sessionId) {
            params.sessionId = _sessionId;
        }
        return params;
    }

    async function call(cmd, params) {
        // IPC-only commands: use Tauri invoke directly
        if (_isIpcCommand(cmd) && _isDesktop) {
            return _invokeTauri(cmd, params);
        }

        var endpoint = _mapCommandToEndpoint(cmd);
        if (!endpoint) {
            // Fallback to Tauri IPC for unmapped commands on desktop
            if (_isDesktop) {
                return _invokeTauri(cmd, params);
            }
            throw new Error('Command not available in Web mode: ' + cmd);
        }

        // Special handling for file upload
        if (cmd === 'open_invoice_files') {
            return _uploadFiles(params);
        }

        // Special handling for health check (returns different format)
        if (cmd === 'check_pdfium_available') {
            var resp = await _fetch(endpoint, {});
            return resp.data && resp.data.pdfium === true;
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
        // Desktop: use Tauri file dialog + server open_invoice_files
        if (_isDesktop && params && params.paths) {
            var body = _addSessionId({ paths: params.paths });
            var result = await _fetch('/api/v1/upload', body);
            if (result.ok && result.data) {
                _sessionId = result.data.sessionId || _sessionId;
                return result.data.files || result.data;
            }
            return result.data;
        }

        // Web: use multipart upload
        // This is handled by the drag/drop or file input handler
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

    return {
        call: call,
        uploadFilesFromInput: uploadFilesFromInput,
        setServerPort: setServerPort,
        setSessionId: setSessionId,
        getSessionId: getSessionId,
        isDesktop: isDesktop,
        getBaseUrl: getBaseUrl,
    };
})();
