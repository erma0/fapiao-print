/**
 * 发票酱 — Web API 层（纯 HTTP）
 */
var __api = (function () {
    var _baseUrl = window.__apiBaseUrl || 'http://127.0.0.1:3000';
    var _sessionId = '';

    var _endpoints = {
        'render_pdf_pages': '/api/v1/render_pdf',
        'render_pdf_pages_pdfium': '/api/v1/render_pdf',
        'extract_pdf_text': '/api/v1/extract_pdf_text',
        'extract_pdf_texts': '/api/v1/extract_pdf_texts',
        'generate_pdf_from_layout': '/api/v1/generate_pdf',
        'get_printers': '/api/v1/printers',
        'list_printers': '/api/v1/printers',
        'pdfium_print': '/api/v1/pdfium_print',
        'print_pdf_file': '/api/v1/print',
        'parse_ofd': '/api/v1/parse_ofd',
        'parse_xml_invoice': '/api/v1/parse_xml_invoice',
        'open_ofd_images': '/api/v1/open_ofd_images',
        'check_path_exists': '/api/v1/check_path_exists',
        'get_config': '/api/v1/get_config',
        'get_app_version': '/api/v1/get_app_version',
        'trim_image': '/api/v1/trim_image',
        'copy_file': '/api/v1/copy_file',
        'rename_file': '/api/v1/rename_file',
        'write_text_file': '/api/v1/write_text_file',
        'get_temp_dir': '/api/v1/get_temp_dir',
        'get_downloads_dir': '/api/v1/get_downloads_dir',
    };

    function _mapEndpoint(cmd) {
        return _endpoints[cmd] || null;
    }

    async function call(cmd, params) {
        if (cmd === 'check_pdfium_available' || cmd === 'check_ocr_available') {
            var resp = await _fetch('/api/v1/health', {}, 'GET');
            return resp.data && (cmd === 'check_ocr_available' ? resp.data.ocr === true : resp.data.pdfium === true);
        }

        var endpoint = _mapEndpoint(cmd);
        if (!endpoint) {
            throw new Error('Command not available: ' + cmd);
        }

        var body = Object.assign({ sessionId: _sessionId || '' }, params || {});
        var result = await _fetch(endpoint, body, 'POST');

        if (!result.ok) {
            var err = new Error(result.error || 'Unknown error');
            err.code = result.code;
            err.recoverable = result.recoverable;
            throw err;
        }
        return result.data;
    }

    async function _fetch(endpoint, body, method) {
        var url = _baseUrl + endpoint;
        var opts = {
            method: method || 'POST',
            headers: method === 'GET' ? {} : { 'Content-Type': 'application/json' },
        };
        if (method !== 'GET') opts.body = JSON.stringify(body);

        var response = await fetch(url, opts);
        if (!response.ok) {
            var errorBody;
            try { errorBody = await response.json(); } catch (_) {
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

    async function uploadFilesFromInput(fileList) {
        var formData = new FormData();
        for (var i = 0; i < fileList.length; i++) {
            formData.append('files', fileList[i]);
        }
        var response = await fetch(_baseUrl + '/api/v1/upload', { method: 'POST', body: formData });
        if (!response.ok) throw new Error('Upload failed: HTTP ' + response.status);
        var result = await response.json();
        if (result.ok && result.data) {
            _sessionId = result.data.sessionId || _sessionId;
            return result.data.files;
        }
        throw new Error(result.error || 'Upload failed');
    }

    function listen(taskId, callback) {
        var source = new EventSource(_baseUrl + '/api/v1/progress/' + taskId);
        source.onmessage = function(e) {
            try { callback(JSON.parse(e.data)); } catch(ex) { callback(e.data); }
        };
        source.onerror = function() { source.close(); };
        return { close: function() { source.close(); } };
    }

    function downloadFile(sessionId, filename) {
        var a = document.createElement('a');
        a.href = _baseUrl + '/api/v1/download/' + sessionId + '/' + encodeURIComponent(filename);
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    function openUrl(url) {
        window.open(url, '_blank');
    }

    function generateTaskId() {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
            var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    function getSessionId() { return _sessionId; }
    function setSessionId(id) { _sessionId = id; }
    function getBaseUrl() { return _baseUrl; }

    async function init() {
        try {
            var resp = await fetch(_baseUrl + '/api/v1/health');
            return resp.ok;
        } catch(e) {
            console.warn('Server not reachable at ' + _baseUrl);
            return false;
        }
    }

    return {
        init: init,
        call: call,
        listen: listen,
        generateTaskId: generateTaskId,
        uploadFilesFromInput: uploadFilesFromInput,
        downloadFile: downloadFile,
        openUrl: openUrl,
        setSessionId: setSessionId,
        getSessionId: getSessionId,
        getBaseUrl: getBaseUrl,
    };
})();
