function FileSaveAs(filename, fileContent, fileType = "text/plain") {
    let link = document.createElement('a');
    link.download = filename;
    link.href = `data:${fileType};base64,${fileContent}`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}