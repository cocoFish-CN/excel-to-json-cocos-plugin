"use strict";
/**
 * Excel 转换插件主入口
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
exports.load = function () {
    console.log('[Excel Plugin] Plugin loaded!');
};
exports.unload = function () {
    console.log('[Excel Plugin] Plugin unloaded!');
};
exports.methods = {
    'open-panel': function () {
        console.log('[Excel Plugin] open-panel method called!');
        try {
            // @ts-ignore
            if (Editor && Editor.Panel) {
                // @ts-ignore
                Editor.Panel.open('excel-to-json.excel-panel');
                console.log('[Excel Plugin] Panel opened');
            }
            else {
                console.error('[Excel Plugin] Editor.Panel not available');
            }
        }
        catch (error) {
            console.error('[Excel Plugin] Error opening panel:', error);
        }
    },
    'set-output-path': function (path) {
        console.log('[Excel Plugin] Setting output path:', path);
    },
    'convert-files': function (data) {
        console.log('[Excel Plugin] Converting files:', data);
        console.log('[Excel Plugin] Files:', data.files);
        console.log('[Excel Plugin] Output path:', data.outputPath);
        console.log('[Excel Plugin] Generate types:', data.generateTypes);
        try {
            convertExcelFiles(data.files, data.outputPath, data.generateTypes);
        }
        catch (error) {
            console.error('[Excel Plugin] Error converting files:', error);
            // @ts-ignore
            Editor.Dialog.error('转换失败', error.message);
        }
    },
    'open-output-folder': function (outputPath) {
        console.log('[Excel Plugin] Opening output folder:', outputPath);
        try {
            const { exec } = require('child_process');
            const resolvedPath = path.resolve(outputPath);
            // 检查路径是否存在
            if (!fs.existsSync(resolvedPath)) {
                // @ts-ignore
                Editor.Dialog.warn('文件夹不存在', `路径 "${resolvedPath}" 不存在`);
                return;
            }
            // 在Windows上使用explorer命令打开文件夹
            if (process.platform === 'win32') {
                // Windows explorer命令，使用不同的处理方式
                exec(`explorer "${resolvedPath}"`, { timeout: 5000 }, (error) => {
                    // Windows explorer经常返回非零退出码但实际成功打开
                    // 只有在真正无法执行命令时才显示错误
                    if (error && error.code === 'ENOENT') {
                        console.error('[Excel Plugin] Explorer not found:', error);
                        // @ts-ignore
                        Editor.Dialog.error('打开文件夹失败', '无法找到Windows资源管理器');
                    }
                    else {
                        console.log('[Excel Plugin] Folder opened successfully with explorer');
                    }
                });
            }
            else {
                // Mac和Linux系统
                const command = process.platform === 'darwin' ?
                    `open "${resolvedPath}"` :
                    `xdg-open "${resolvedPath}"`;
                exec(command, (error) => {
                    if (error) {
                        console.error('[Excel Plugin] Error opening folder:', error);
                        // @ts-ignore
                        Editor.Dialog.error('打开文件夹失败', error.message);
                    }
                    else {
                        console.log('[Excel Plugin] Folder opened successfully');
                    }
                });
            }
        }
        catch (error) {
            console.error('[Excel Plugin] Error opening folder:', error);
            // @ts-ignore
            Editor.Dialog.error('打开文件夹失败', error.message);
        }
    }
};
function convertExcelFiles(files, outputPath, generateTypes) {
    files.forEach(file => {
        try {
            console.log('[Excel Plugin] Processing file:', file);
            const workbook = XLSX.readFile(file);
            const baseName = path.basename(file, path.extname(file));
            // 处理多工作表情况
            const hasMultipleSheets = workbook.SheetNames.length > 1;
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const rawData = XLSX.utils.sheet_to_json(worksheet);
                // 智能处理Excel数据，转换为键值对对象格式
                const processedData = processExcelData(rawData);
                // 确定输出文件名
                const jsonFileName = hasMultipleSheets ?
                    `${baseName}_${sheetName}.json` :
                    `${baseName}.json`;
                const jsonFilePath = path.join(outputPath, jsonFileName);
                // 确保输出目录存在
                fs.mkdirSync(path.dirname(jsonFilePath), { recursive: true });
                // 写入JSON文件
                fs.writeFileSync(jsonFilePath, JSON.stringify(processedData, null, 2), 'utf8');
                console.log('[Excel Plugin] Generated JSON file:', jsonFilePath);
                // 生成TypeScript类型定义
                if (generateTypes && Object.keys(processedData).length > 0) {
                    generateTypeDefinitionForObject(processedData, outputPath, baseName, sheetName, hasMultipleSheets);
                }
            });
            // @ts-ignore
            Editor.Dialog.info('转换完成', `成功转换文件: ${file}`);
        }
        catch (error) {
            console.error('[Excel Plugin] Error processing file:', file, error);
            // @ts-ignore
            Editor.Dialog.error('文件转换失败', `处理文件 ${file} 时出错: ${error.message}`);
        }
    });
}
function generateTypeDefinition(jsonData, outputPath, baseName, sheetName, hasMultipleSheets) {
    try {
        const sampleRow = jsonData[0];
        const interfaceName = hasMultipleSheets ?
            `${toPascalCase(baseName)}${toPascalCase(sheetName)}` :
            toPascalCase(baseName);
        let typeContent = `// Auto-generated TypeScript interface\n\n`;
        typeContent += `export interface ${interfaceName} {\n`;
        Object.keys(sampleRow).forEach(key => {
            const value = sampleRow[key];
            let type = 'any';
            if (typeof value === 'string') {
                type = 'string';
            }
            else if (typeof value === 'number') {
                type = 'number';
            }
            else if (typeof value === 'boolean') {
                type = 'boolean';
            }
            else if (value instanceof Date) {
                type = 'Date';
            }
            // 智能处理属性名：保持有效标识符，转换中文为拼音或使用引号
            const cleanKey = generateValidPropertyName(key);
            typeContent += `  ${cleanKey}: ${type};\n`;
        });
        typeContent += `}\n\n`;
        typeContent += `export type ${interfaceName}Array = ${interfaceName}[];\n`;
        // 确定输出文件名
        const typeFileName = hasMultipleSheets ?
            `${baseName}_${sheetName}.d.ts` :
            `${baseName}.d.ts`;
        const typeFilePath = path.join(outputPath, typeFileName);
        fs.writeFileSync(typeFilePath, typeContent, 'utf8');
        console.log('[Excel Plugin] Generated TypeScript types:', typeFilePath);
    }
    catch (typeError) {
        console.error('[Excel Plugin] Error generating types:', typeError);
    }
}
function generateValidPropertyName(key) {
    // 如果包含中文或特殊字符，使用引号包围
    if (/[^\w$]/.test(key)) {
        return `"${key}"`;
    }
    // 如果以数字开头，添加引号
    if (/^\d/.test(key)) {
        return `"${key}"`;
    }
    // 如果是纯英文数字下划线，直接使用
    return key;
}
function toPascalCase(str) {
    return str.replace(/[^a-zA-Z0-9]/g, ' ')
        .replace(/\b\w/g, l => l.toUpperCase())
        .replace(/\s/g, '');
}
/**
 * 处理Excel数据，将数组转换为键值对对象
 * 过滤配置行，只保留实际数据行
 */
function processExcelData(rawData) {
    if (!rawData || rawData.length === 0) {
        return {};
    }
    console.log('[Excel Plugin] Processing raw data:', rawData.length, 'rows');
    // 找到第一行（字段映射行）
    const headerRow = rawData[0];
    if (!headerRow) {
        return {};
    }
    // 找到KEY字段名
    const keyFieldName = Object.keys(headerRow).find(key => key.includes('KEY') || key.includes('key') || key.includes('Key') ||
        key.includes('编号') || key.includes('ID') || key.includes('id'));
    if (!keyFieldName) {
        console.warn('[Excel Plugin] No KEY field found in:', Object.keys(headerRow));
        return {};
    }
    console.log('[Excel Plugin] Using KEY field:', keyFieldName);
    // 获取字段映射（第一行）
    const fieldMapping = headerRow;
    // 过滤数据行：跳过前面的配置行，只处理实际数据
    const dataRows = rawData.filter((row, index) => {
        if (index < 4)
            return false; // 跳过前4行（字段映射、类型、server、client）
        const keyValue = row[keyFieldName];
        // 检查KEY值是否为有效的数据ID（数字）
        return keyValue !== undefined && keyValue !== null &&
            !isNaN(Number(keyValue)) && keyValue !== '';
    });
    console.log('[Excel Plugin] Found data rows:', dataRows.length);
    const result = {};
    dataRows.forEach(row => {
        const keyValue = row[keyFieldName];
        if (keyValue === undefined || keyValue === null)
            return;
        const dataObject = {};
        // 根据字段映射转换数据
        Object.keys(fieldMapping).forEach(originalField => {
            const mappedField = fieldMapping[originalField];
            const value = row[originalField];
            // 跳过KEY字段，因为它已经作为对象的键
            if (originalField === keyFieldName)
                return;
            // 跳过空的映射字段
            if (!mappedField || mappedField === '')
                return;
            // 添加到数据对象
            if (value !== undefined && value !== null && value !== '') {
                dataObject[mappedField] = value;
            }
        });
        result[keyValue] = dataObject;
    });
    console.log('[Excel Plugin] Generated object with keys:', Object.keys(result));
    return result;
}
/**
 * 为对象格式数据生成TypeScript类型定义
 */
function generateTypeDefinitionForObject(data, outputPath, baseName, sheetName, hasMultipleSheets) {
    try {
        const keys = Object.keys(data);
        if (keys.length === 0)
            return;
        const sampleObject = data[keys[0]];
        const interfaceName = hasMultipleSheets ?
            `${toPascalCase(baseName)}${toPascalCase(sheetName)}` :
            toPascalCase(baseName);
        let typeContent = `// Auto-generated TypeScript interface for object-format data\n\n`;
        // 数据项接口
        typeContent += `export interface ${interfaceName}Item {\n`;
        Object.keys(sampleObject).forEach(key => {
            const value = sampleObject[key];
            let type = 'any';
            if (typeof value === 'string') {
                type = 'string';
            }
            else if (typeof value === 'number') {
                type = 'number';
            }
            else if (typeof value === 'boolean') {
                type = 'boolean';
            }
            else if (value instanceof Date) {
                type = 'Date';
            }
            const cleanKey = generateValidPropertyName(key);
            typeContent += `  ${cleanKey}: ${type};\n`;
        });
        typeContent += `}\n\n`;
        // 整个数据集合接口
        typeContent += `export interface ${interfaceName}Data {\n`;
        typeContent += `  [key: string]: ${interfaceName}Item;\n`;
        typeContent += `}\n\n`;
        // 类型别名
        typeContent += `export type ${interfaceName} = ${interfaceName}Data;\n`;
        // 确定输出文件名
        const typeFileName = hasMultipleSheets ?
            `${baseName}_${sheetName}.d.ts` :
            `${baseName}.d.ts`;
        const typeFilePath = path.join(outputPath, typeFileName);
        fs.writeFileSync(typeFilePath, typeContent, 'utf8');
        console.log('[Excel Plugin] Generated TypeScript types for object format:', typeFilePath);
    }
    catch (typeError) {
        console.error('[Excel Plugin] Error generating object types:', typeError);
    }
}
