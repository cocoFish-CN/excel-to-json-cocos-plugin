/**
 * Excel 转换面板
 */

// @ts-ignore
const Editor = globalThis.Editor;
const pathUtil = require('path');

exports.template = `
<div class="excel-converter-panel">
    <ui-section>
        <div class="header">
            <h1>Excel 转换工具</h1>
        </div>
        <ui-prop name="输出目录">
            <ui-input class="output-path" value="assets/resources"></ui-input>
            <ui-button class="select-output">选择目录</ui-button>
        </ui-prop>
        <ui-prop name="Excel文件">
            <ui-button class="select-files">选择Excel文件</ui-button>
            <ui-button class="convert-files" style="margin-left: 10px;">开始转换</ui-button>
        </ui-prop>
        <ui-prop name="选项">
            <ui-checkbox class="generate-types">生成TypeScript类型定义</ui-checkbox>
        </ui-prop>
        <div class="file-list">
            <div class="files-container"></div>
        </div>
    </ui-section>
</div>
`;

exports.style = `
.excel-converter-panel {
    padding: 10px;
}

.header {
    margin-bottom: 20px;
}

.header h1 {
    color: var(--color-normal-contrast-weakest);
    font-size: 16px;
}

.file-list {
    margin-top: 20px;
    border: 1px solid var(--color-normal-border);
    padding: 10px;
    min-height: 100px;
    max-height: 300px;
    overflow-y: auto;
}

.files-container {
    display: flex;
    flex-direction: column;
    gap: 5px;
}

.file-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 8px;
    background: var(--color-normal-fill-emphasis);
    border-radius: 4px;
}

.file-name {
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.remove-btn {
    background: var(--color-warn-fill);
    color: white;
    border-radius: 50%;
    width: 20px;
    height: 20px;
    display: flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
    font-size: 14px;
    margin-left: 10px;
}

.remove-btn:hover {
    background: var(--color-warn-fill-emphasis);
}

.empty-message {
    text-align: center;
    color: var(--color-normal-contrast-weaker);
    font-style: italic;
    padding: 20px 0;
}
`;

exports.$ = {
    outputPath: '.output-path',
    selectOutput: '.select-output',
    selectFiles: '.select-files',
    convertFiles: '.convert-files',
    filesContainer: '.files-container',
    generateTypes: '.generate-types'
};

let fileList: string[] = [];
let lastOutputPath: string = ''; // 记录上次选择的输出目录
let lastFilePath: string = ''; // 记录上次选择文件的目录

exports.ready = function () {
    console.log('[Excel Panel] Panel ready!');

    // 保存 this 的引用
    const self = this;

    // 绑定 updateFileList 到 self
    self.updateFileList = exports.updateFileList.bind(self);

    // 默认选中生成类型定义
    this.$.generateTypes.checked = true;

    // 选择输出目录
    this.$.selectOutput.addEventListener('click', async () => {
        try {
            // 使用记忆的路径或项目根目录
            const defaultPath = lastOutputPath ||
                (self.$.outputPath.value.trim() && self.$.outputPath.value.trim() !== 'assets/resources' ?
                    self.$.outputPath.value.trim() :
                    // @ts-ignore
                    Editor.Project.path);

            // @ts-ignore
            const result = await Editor.Dialog.select({
                title: '选择输出目录',
                type: 'directory',
                path: defaultPath
            });

            if (!result.canceled && result.filePaths.length > 0) {
                const selectedPath = result.filePaths[0];
                self.$.outputPath.value = selectedPath;
                lastOutputPath = selectedPath; // 记录这次选择的路径
                console.log('[Excel Panel] Remembered output path:', lastOutputPath);

                // @ts-ignore
                Editor.Message.send('excel-to-json', 'set-output-path', selectedPath);
            }
        } catch (err) {
            console.error('[Excel Panel] Select output directory failed:', err);
            // @ts-ignore
            Editor.Dialog.warn('选择目录失败', err.message);
        }
    });

    // 选择Excel文件
    this.$.selectFiles.addEventListener('click', async () => {
        try {
            // 使用记忆的路径或项目根目录
            const defaultPath = lastFilePath ||
                // @ts-ignore
                Editor.Project.path;

            console.log('[Excel Panel] Using path for file selection:', defaultPath);

            // @ts-ignore
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
                lastFilePath = pathUtil.dirname(result.filePaths[0]);
                console.log('[Excel Panel] Remembered file path:', lastFilePath);

                // 添加新文件到列表，避免重复
                result.filePaths.forEach((filePath: string) => {
                    if (!fileList.includes(filePath)) {
                        fileList.push(filePath);
                    }
                });
                self.updateFileList();
            }
        } catch (err) {
            console.error('[Excel Panel] Select files failed:', err);
            // @ts-ignore
            Editor.Dialog.warn('选择文件失败', err.message);
        }
    });

    // 开始转换
    this.$.convertFiles.addEventListener('click', () => {
        try {
            if (fileList.length === 0) {
                // @ts-ignore
                Editor.Dialog.warn('请先选择Excel文件！');
                return;
            }

            const outputPath = self.$.outputPath.value.trim();
            if (!outputPath) {
                // @ts-ignore
                Editor.Dialog.warn('请设置输出目录！');
                return;
            }

            console.log('[Excel Panel] Starting conversion...');
            // @ts-ignore
            Editor.Message.send('excel-to-json', 'convert-files', {
                files: fileList,
                outputPath: outputPath,
                generateTypes: self.$.generateTypes.checked
            });

            console.log('[Excel Panel] Conversion request sent to main process');
        } catch (err) {
            console.error('[Excel Panel] Convert files failed:', err);
            // @ts-ignore
            Editor.Dialog.warn('转换失败', err.message);
        }
    });
};

exports.updateFileList = function () {
    const self = this;

    if (!this.$.filesContainer) return;

    this.$.filesContainer.innerHTML = '';

    if (fileList.length === 0) {
        const emptyDiv = document.createElement('div');
        emptyDiv.className = 'empty-message';
        emptyDiv.textContent = '尚未选择任何文件';
        self.$.filesContainer.appendChild(emptyDiv);
        return;
    }

    fileList.forEach((file, index) => {
        const div = document.createElement('div');
        div.className = 'file-item';

        const fileName = document.createElement('span');
        fileName.className = 'file-name';
        fileName.textContent = file;

        const removeBtn = document.createElement('span');
        removeBtn.className = 'remove-btn';
        removeBtn.textContent = '×';
        removeBtn.title = '移除此文件';
        removeBtn.addEventListener('click', () => {
            fileList.splice(index, 1);
            self.updateFileList();
        });

        div.appendChild(fileName);
        div.appendChild(removeBtn);
        self.$.filesContainer.appendChild(div);
    });
};
