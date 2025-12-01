// Business Card Maker - JavaScript Helper Functions

/**
 * Downloads a file from base64 data
 * @param {string} filename - The name of the file to download
 * @param {string} base64Data - The base64 encoded file data
 */
window.downloadFile = function (filename, base64Data) {
    const blob = base64ToBlob(base64Data);
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
};

/**
 * Converts base64 string to Blob
 * @param {string} base64 - The base64 encoded data
 * @returns {Blob} - The blob object
 */
function base64ToBlob(base64) {
    const binaryString = window.atob(base64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return new Blob([bytes], { type: 'application/octet-stream' });
}

/**
 * Downloads a file from a Blazor stream reference (used for PPTX/Excel template downloads).
 * @param {string} filename - The name of the file to download.
 * @param {any} contentStreamReference - The DotNetStreamReference from .NET.
 */
window.downloadFileFromStream = async (filename, contentStreamReference) => {
    const arrayBuffer = await contentStreamReference.arrayBuffer();
    const blob = new Blob([arrayBuffer], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const anchorElement = document.createElement('a');
    anchorElement.href = url;
    anchorElement.download = filename ?? 'download';
    anchorElement.click();
    anchorElement.remove();
    URL.revokeObjectURL(url);
};
