const fs = require('fs');
const path = require('path');
const obfuscator = require('javascript-obfuscator');

const scanDirectoryPath = '../FractalExcelWebAddInWeb';
const outputDir = './dist';

function scanDir(scanPath, codeExtentions, filters) {
    if (!fs.existsSync(scanPath)) {
        return [];
    }

    const checkFilters = (filePath) => {
        const includePath = filters.codeIncludePath || [];
        const excludePath = filters.codeExcludePath || [];
        const fullIgnorePath = filters.codeFullIgnore || [];
        const staticPath = filters.staticPath || [];

        const extention = codeExtentions.some(ext => path.extname(filePath) === ext);
        const include = includePath.some(filter => filter.test(filePath));
        const exclude = excludePath.some(filter => filter.test(filePath));
        const fullIgnore = fullIgnorePath.some(filter => filter.test(filePath));
        const static = staticPath.some(filter => filter.test(filePath));

        return (extention && (include || !exclude) || static) && !fullIgnore;
    }

    const result = [];
    const files = fs.readdirSync(scanPath);

    files.forEach(file => {
        const filePath = path.join(scanPath, file);
        const stat = fs.lstatSync(filePath);

        if (stat.isDirectory()) {
            result.push.apply(result, scanDir(filePath, codeExtentions, filters));
        } else if (checkFilters(filePath)) {
            result.push(filePath);
        }
    });

    return result;
}

function build() {
    const files = scanDir(
        scanDirectoryPath,
        ['.css', '.html', '.js', ],
        {
            codeIncludePath: [/Scripts\/Login.js/],
            codeExcludePath: [/Content/, /Scripts/],
            codeFullIgnore: [/bin/, /obj/, /node_modules/],

            staticPath: [/Images/]
        }
    );

    files.forEach(file => {
        const localPath = file.slice(scanDirectoryPath.length + 1);
        const dstPath = path.join(outputDir, localPath);

        fs.mkdirSync(path.dirname(dstPath), { recursive: true });

        if(path.extname(file) === '.js') {
            // Obfuscate
            const fileContents = fs.readFileSync(file);
            const obfuscatedCode = obfuscator.obfuscate(fileContents.toString());

            fs.writeFileSync(dstPath, obfuscatedCode.getObfuscatedCode());
        } else {
            // Copy to dist
            fs.copyFileSync(file, dstPath);
        }
    })
}

build();
