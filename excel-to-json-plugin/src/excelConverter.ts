import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

interface ConverterOptions {
    generateTypes?: boolean;
    outputDir?: string;
}

export class ExcelConverter {
    constructor(private options: ConverterOptions = {}) {
        this.options.outputDir = this.options.outputDir || 'assets/resources';
        this.options.generateTypes = this.options.generateTypes !== false;
    }

    /**
     * 将 Excel 文件转换为 JSON 和 TypeScript 类型定义
     * @param filePath Excel 文件路径
     */
    public async convert(filePath: string): Promise<void> {
        // 读取 Excel 文件
        const workbook = XLSX.readFile(filePath);
        
        // 确保输出目录存在
        const fullOutputPath = path.join(Editor.Project.path, this.options.outputDir!);
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
    private generateTypeDefinition(data: any, interfaceName: string): string {
        const properties = Object.entries(data).map(([key, value]) => {
            const type = this.getTypeScriptType(value);
            return `  ${this.sanitizePropertyName(key)}: ${type};`;
        });

        return `export interface ${this.sanitizeInterfaceName(interfaceName)} {\n${properties.join('\n')}\n}\n`;
    }

    /**
     * 获取值的 TypeScript 类型
     */
    private getTypeScriptType(value: any): string {
        if (value === null) return 'null';
        if (Array.isArray(value)) return 'any[]';
        
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
    private sanitizeInterfaceName(name: string): string {
        const sanitized = name
            .replace(/[^a-zA-Z0-9]/g, '_')
            .replace(/^[0-9]/, '_$&');
        return sanitized.charAt(0).toUpperCase() + sanitized.slice(1) + 'Data';
    }

    /**
     * 清理属性名称
     */
    private sanitizePropertyName(name: string): string {
        return name.replace(/[【】（）\s]/g, '')
            .replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '_');
    }
}
