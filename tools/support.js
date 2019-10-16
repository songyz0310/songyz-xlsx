// 将指定的自然数转换为26进制表示。映射关系：[0-25] -> [A-Z]。
export const getCharCol = (n) => {
    let s = '',
        m = 0
    while (n > 0) {
        m = n % 26 + 1
        s = String.fromCharCode(m + 64) + s
        n = (n - m) / 26
    }
    return s
}

//将数据写到文件中
export const writeFile = (fname, data, enc) => {
    /*global IE_SaveFile, Blob, navigator, saveAs, URL, document, File, chrome */
    if (typeof IE_SaveFile !== 'undefined') return IE_SaveFile(data, fname);
    if (typeof Blob !== 'undefined') {
        var blob = new Blob([blobify(data)], { type: "application/octet-stream" });
        if (typeof navigator !== 'undefined' && navigator.msSaveBlob) return navigator.msSaveBlob(blob, fname);
        if (typeof saveAs !== 'undefined') return saveAs(blob, fname);
        if (typeof URL !== 'undefined' && typeof document !== 'undefined' && document.createElement && URL.createObjectURL) {
            var url = URL.createObjectURL(blob);
            if (typeof chrome === 'object' && typeof(chrome.downloads || {}).download == "function") {
                if (URL.revokeObjectURL && typeof setTimeout !== 'undefined') setTimeout(function() { URL.revokeObjectURL(url); }, 60000);
                return chrome.downloads.download({ url: url, filename: fname, saveAs: true });
            }
            var a = document.createElement("a");
            if (a.download != null) {
                a.download = fname;
                a.href = url;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                if (URL.revokeObjectURL && typeof setTimeout !== 'undefined') setTimeout(function() { URL.revokeObjectURL(url); }, 60000);
                return url;
            }
        }
    }
    // $FlowIgnore
    if (typeof $ !== 'undefined' && typeof File !== 'undefined' && typeof Folder !== 'undefined') try { // extendscript
        // $FlowIgnore
        var out = File(fname);
        out.open("w");
        out.encoding = "binary";
        if (Array.isArray(payload)) payload = a2s(payload);
        out.write(payload);
        out.close();
        return payload;
    } catch (e) { if (!e.message || !e.message.match(/onstruct/)) throw e; }
    throw new Error("cannot save file " + fname);
}

/* normalize data for blob ctor */
function blobify(data) {
    if (typeof data === "string") return s2ab(data);
    if (Array.isArray(data)) return a2u(data);
    return data;
}

function s2ab(s) {
    if (typeof ArrayBuffer === 'undefined') return s2a(s);
    var buf = new ArrayBuffer(s.length),
        view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

function a2u(data) {
    if (typeof Uint8Array === 'undefined') throw new Error("Unsupported");
    return new Uint8Array(data);
}