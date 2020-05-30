export function isXlsx(fileName: string){
    const parts = fileName.split('.');
    const suffix = parts[parts.length - 1];
    return suffix.toLowerCase() === "xlsx";
}