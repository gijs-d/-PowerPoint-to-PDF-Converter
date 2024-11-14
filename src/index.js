const path = require('path');
const fs = require('fs').promises;

const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);

const inputPath = path.join(__dirname, '../convert/in-pptx');
const outputPath = path.join(__dirname, '../convert/out-pdf');

(async () => {
    try {
        await clearOutDir();
        await convertFiles();
    } catch (e) {
        console.error(`Error converting file: ${e}`);
    }
})();

async function convertFiles() {
    const ext = '.pdf';
    const files = await fs.readdir(inputPath);
    console.log(`start converting ${files.length} office documents:`);
    files.forEach(async (file, i) => {
        console.log(`\t in -  ${i} - ${file}`);
        const filePath = path.join(inputPath, '/' + file);
        const pptxBuffer = await fs.readFile(filePath);
        const pdfBuf = await libre.convertAsync(pptxBuffer, ext, undefined);
        const fileName = `${file.split('.')[0]}${ext}`;
        await fs.writeFile(path.join(outputPath, `/${fileName}`), pdfBuf);
        console.log(`\t out -  ${i} - ${fileName}`);
    });
}

async function clearOutDir() {
    const files = await fs.readdir(outputPath);
    console.log(`clear ${files.length} old outputs:`);
    files.forEach(async (f, file) => {
        console.log('\t clear ', file, f);
        await fs.unlink(path.join(outputPath, `/${f}`));
    });
    console.log();
}