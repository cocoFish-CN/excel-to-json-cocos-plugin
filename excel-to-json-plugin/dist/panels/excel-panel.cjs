// Excel转换工具面板 - 路径记忆版本
const Editor = globalThis.Editor;

exports.template = `
<div style="padding: 15px; font-family: 'Segoe UI', Arial, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 8px; margin: 5px; min-height: 95vh; overflow-y: auto;">
    <div style="background: rgba(255,255,255,0.95); padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <h2 style="color: #2c3e50; margin: 0 0 20px 0; text-align: center; font-size: 24px; text-shadow: 2px 2px 4px rgba(0,0,0,0.1);">
            🚀 Excel 转换工具 🚀
        </h2>
        
        <div style="margin-bottom: 15px; padding: 15px; background: #f8f9fa; border-radius: 6px; border-left: 4px solid #007ACC;">
            <label style="display: block; margin-bottom: 8px; font-weight: bold; color: #2c3e50;">📂 输出目录:</label>
            <div style="display: flex; align-items: center; gap: 10px;">
                <input type="text" id="outputPath" value="assets/resources" 
                       style="flex: 1; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                <button id="selectOutput" style="padding: 12px 20px; background: linear-gradient(45deg, #007ACC, #005a9e); color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: bold;">📁 选择目录</button>
            </div>
        </div>
        
        <div style="margin-bottom: 15px; padding: 15px; background: #f8f9fa; border-radius: 6px; border-left: 4px solid #28a745;">
            <div style="display: flex; gap: 15px; flex-wrap: wrap;">
                <button id="selectFiles" style="padding: 12px 20px; background: linear-gradient(45deg, #28a745, #20c997); color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: bold;">📊 选择Excel文件</button>
                <button id="convertFiles" style="padding: 12px 20px; background: linear-gradient(45deg, #dc3545, #c82333); color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: bold;">⚡ 开始转换</button>
            </div>
        </div>
        
        <div style="margin-bottom: 15px; padding: 15px; background: #f8f9fa; border-radius: 6px; border-left: 4px solid #6f42c1;">
            <label style="display: flex; align-items: center; font-weight: bold; color: #2c3e50;">
                <input type="checkbox" id="generateTypes" checked style="margin-right: 10px; transform: scale(1.2);">
                🔧 生成TypeScript类型定义
            </label>
        </div>
        
        <div style="border: 2px solid #ddd; padding: 20px; border-radius: 8px; min-height: 150px; background: #ffffff; box-shadow: inset 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 15px;">
            <h4 style="margin: 0 0 15px 0; color: #2c3e50; font-size: 16px;">📋 选中的文件:</h4>
            <div id="filesContainer" style="max-height: 250px; overflow-y: auto;">
                <div style="text-align: center; color: #6c757d; font-style: italic; padding: 30px 0; font-size: 16px;">
                    📂 尚未选择任何文件
                </div>
            </div>
        </div>
        
        <div style="text-align: center;">
            <button id="openOutputFolder" style="padding: 15px 30px; background: linear-gradient(45deg, #6f42c1, #5a32a3); color: white; border: none; border-radius: 8px; cursor: pointer; font-size: 16px; font-weight: bold;">
                📁 打开转换文件夹 📁
            </button>
            <div style="margin-top: 15px; color: #28a745; font-size: 14px; font-weight: bold;">
                ✅ v6.0 ✅
            </div>
        </div>
    </div>
</div>
`;

exports.style = `
button:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 8px rgba(0,0,0,0.2);
}
input:focus {
    outline: none;
    border-color: #007ACC !important;
    box-shadow: 0 0 0 3px rgba(0, 122, 204, 0.3);
}
.file-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 12px 15px;
    margin: 8px 0;
    background: linear-gradient(45deg, #f8f9fa, #e9ecef);
    border: 1px solid #ddd;
    border-radius: 6px;
    transition: all 0.3s;
}
.file-item:hover {
    background: linear-gradient(45deg, #e9ecef, #dee2e6);
    transform: translateX(5px);
}
.file-name {
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    margin-right: 15px;
    font-weight: 500;
    color: #2c3e50;
}
.remove-btn {
    background: linear-gradient(45deg, #ff4444, #cc0000);
    color: white;
    border: none;
    border-radius: 50%;
    width: 28px;
    height: 28px;
    cursor: pointer;
    font-size: 16px;
    font-weight: bold;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: all 0.3s;
}
.remove-btn:hover {
    background: linear-gradient(45deg, #cc0000, #990000);
    transform: scale(1.1);
}
`;

exports.$ = {
    outputPath: '#outputPath',
    selectOutput: '#selectOutput',
    selectFiles: '#selectFiles',
    convertFiles: '#convertFiles',
    filesContainer: '#filesContainer',
    generateTypes: '#generateTypes',
    openOutputFolder: '#openOutputFolder'
};

let fileList = [];
let lastOutputPath = ''; // 记录上次选择的输出目录
let lastFilePath = ''; // 记录上次选择文件的目录

exports.ready = function () {
    console.log('[Excel Panel v6.0] 对象格式版面板准备就绪!');

    const self = this;

    // 选择输出目录 - 带路径记忆
    this.$.selectOutput.addEventListener('click', async () => {
        try {
            // 使用记忆的路径或当前输出路径或项目根目录
            const currentPath = self.$.outputPath.value.trim();
            const defaultPath = lastOutputPath ||
                (currentPath && currentPath !== 'assets/resources' ? currentPath : Editor.Project.path);

            console.log('[Excel Panel] 使用记忆的输出路径:', defaultPath);

            const result = await Editor.Dialog.select({
                title: '选择输出目录',
                type: 'directory',
                path: defaultPath
            });

            if (!result.canceled && result.filePaths.length > 0) {
                const selectedPath = result.filePaths[0];
                self.$.outputPath.value = selectedPath;
                lastOutputPath = selectedPath; // 记录这次选择的路径
                console.log('[Excel Panel] 记录新的输出目录路径:', lastOutputPath);
                Editor.Message.send('excel-to-json', 'set-output-path', selectedPath);
            }
        } catch (err) {
            console.error('[Excel Panel] Select output directory failed:', err);
            Editor.Dialog.warn('选择目录失败', err.message);
        }
    });

    // 选择Excel文件 - 带路径记忆
    this.$.selectFiles.addEventListener('click', async () => {
        try {
            // 使用记忆的文件路径或项目根目录
            const defaultPath = lastFilePath || Editor.Project.path;

            console.log('[Excel Panel] 使用记忆的文件路径:', defaultPath);

            const result = await Editor.Dialog.select({
                title: '选择Excel文件',
                type: 'file',
                path: defaultPath,
                filters: [
                    { name: 'Excel文件', extensions: ['xlsx', 'xls'] }
                ],
                properties: ['openFile', 'multiSelections']
            });

            if (!result.canceled && result.filePaths.length > 0) {
                // 记录文件所在目录（使用第一个文件的目录）
                const pathModule = require('path');
                lastFilePath = pathModule.dirname(result.filePaths[0]);
                console.log('[Excel Panel] 记录新的文件选择路径:', lastFilePath);

                result.filePaths.forEach(filePath => {
                    if (!fileList.includes(filePath)) {
                        fileList.push(filePath);
                    }
                });
                updateFileList();
            }
        } catch (err) {
            console.error('[Excel Panel] Select files failed:', err);
            Editor.Dialog.warn('选择文件失败', err.message);
        }
    });

    // 开始转换
    this.$.convertFiles.addEventListener('click', () => {
        try {
            if (fileList.length === 0) {
                Editor.Dialog.warn('请先选择Excel文件！');
                return;
            }

            const outputPath = self.$.outputPath.value.trim();
            if (!outputPath) {
                Editor.Dialog.warn('请设置输出目录！');
                return;
            }

            console.log('[Excel Panel] Starting conversion...');
            Editor.Message.send('excel-to-json', 'convert-files', {
                files: fileList,
                outputPath: outputPath,
                generateTypes: self.$.generateTypes.checked
            });
        } catch (err) {
            console.error('[Excel Panel] Convert files failed:', err);
            Editor.Dialog.warn('转换失败', err.message);
        }
    });

    // 打开转换文件夹
    this.$.openOutputFolder.addEventListener('click', () => {
        try {
            const outputPath = self.$.outputPath.value.trim();
            if (!outputPath) {
                Editor.Dialog.warn('请先设置输出目录！');
                return;
            }

            console.log('[Excel Panel] Opening output folder:', outputPath);
            Editor.Message.send('excel-to-json', 'open-output-folder', outputPath);
        } catch (err) {
            console.error('[Excel Panel] Open output folder failed:', err);
            Editor.Dialog.warn('打开文件夹失败', err.message);
        }
    });

    function updateFileList() {
        if (fileList.length === 0) {
            self.$.filesContainer.innerHTML = '<div style="text-align: center; color: #6c757d; font-style: italic; padding: 30px 0; font-size: 16px;">📂 尚未选择任何文件</div>';
            return;
        }

        let html = '';
        fileList.forEach((file, index) => {
            html += `
                <div class="file-item">
                    <span class="file-name" title="${file}">📄 ${file}</span>
                    <button class="remove-btn" onclick="removeFile(${index})" title="移除此文件">×</button>
                </div>
            `;
        });
        self.$.filesContainer.innerHTML = html;
    }

    // 全局函数用于移除文件
    window.removeFile = function (index) {
        fileList.splice(index, 1);
        updateFileList();
    };
};
