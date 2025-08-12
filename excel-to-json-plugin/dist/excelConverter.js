"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelConverter = void 0;
const XLSX = __importStar(require("xlsx"));
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
class ExcelConverter {
    constructor(options = {}) {
        this.options = options;
        this.options.outputDir = this.options.outputDir || 'assets/resources';
        this.options.generateTypes = this.options.generateTypes !== false;
    }
    /**
     * 将 Excel 文件转换为 JSON 和 TypeScript 类型定义
     * @param filePath Excel 文件路径
     */
    async convert(filePath) {
        // 读取 Excel 文件
        const workbook = XLSX.readFile(filePath);
        // 确保输出目录存在
        const fullOutputPath = path.join(Editor.Project.path, this.options.outputDir);
        if (!fs.existsSync(fullOutputPath)) {
            fs.mkdirSync(fullOutputPath, { recursive: true });
        }
        // 获取原始文件名（不含扩展名）
        const originalFileName = path.basename(filePath, path.extname(filePath));
        // 处理每个工作表
        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            if (jsonData.length > 0) {
                // 使用原始文件名生成 JSON 文件
                const jsonFilePath = path.join(fullOutputPath, `${originalFileName}.json`);
                fs.writeFileSync(jsonFilePath, JSON.stringify(jsonData, null, 2));
                // 如果需要生成类型定义
                if (this.options.generateTypes) {
                    const types = this.generateTypeDefinition(jsonData[0], originalFileName);
                    const tsFilePath = path.join(fullOutputPath, `${originalFileName}.ts`);
                    fs.writeFileSync(tsFilePath, types);
                }
            }
        });
        // 刷新资源管理器
        Editor.Message.request('asset-db', 'refresh-asset', `db:///${this.options.outputDir}`);
    }
    /**
     * 根据数据生成 TypeScript 类型定义
     */
    generateTypeDefinition(data, interfaceName) {
        const properties = Object.entries(data).map(([key, value]) => {
            const type = this.getTypeScriptType(value);
            return `  ${this.sanitizePropertyName(key)}: ${type};`;
        });
        return `export interface ${this.sanitizeInterfaceName(interfaceName)} {\n${properties.join('\n')}\n}\n`;
    }
    /**
     * 获取值的 TypeScript 类型
     */
    getTypeScriptType(value) {
        if (value === null)
            return 'null';
        if (Array.isArray(value))
            return 'any[]';
        const type = typeof value;
        switch (type) {
            case 'string':
                return 'string';
            case 'number':
                return Number.isInteger(value) ? 'number' : 'number';
            case 'boolean':
                return 'boolean';
            case 'object':
                return 'Record<string, any>';
            default:
                return 'any';
        }
    }
    /**
     * 清理接口名称
     */
    sanitizeInterfaceName(name) {
        const sanitized = name
            .replace(/[^a-zA-Z0-9]/g, '_')
            .replace(/^[0-9]/, '_$&');
        return sanitized.charAt(0).toUpperCase() + sanitized.slice(1) + 'Data';
    }
    /**
     * 清理属性名称
     */
    sanitizePropertyName(name) {
        return name.replace(/[【】（）\s]/g, '')
            .replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '_');
    }
}
exports.ExcelConverter = ExcelConverter;
