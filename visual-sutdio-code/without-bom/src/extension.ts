import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';

function hasUtf8Bom(buffer: Buffer): boolean {
    return buffer.length >= 3 &&
        buffer[0] === 0xEF &&
        buffer[1] === 0xBB &&
        buffer[2] === 0xBF;
}

function isExcluded(filePath: string, workspaceFolder: string, excluded: string[]): boolean {
    const relativePath = path.relative(workspaceFolder, filePath);
    return excluded.some(folder =>
        relativePath === folder || relativePath.startsWith(folder + path.sep)
    );
}

export function activate(context: vscode.ExtensionContext) {
    const disposable = vscode.workspace.onDidOpenTextDocument(async (document) => {
        if (document.isUntitled || document.languageId === 'binary' || document.isDirty) {
            return;
        }

        const filePath = document.uri.fsPath;
        if (!fs.existsSync(filePath)) {
            return;
        }

        const workspaceFolder = vscode.workspace.getWorkspaceFolder(document.uri)?.uri.fsPath;
        if (!workspaceFolder) {
            return;
        }

        const config = vscode.workspace.getConfiguration('withoutBom');
        const excludedFolders: string[] = config.get('excludeFolders') || [];

        if (isExcluded(filePath, workspaceFolder, excludedFolders)) {
            return;
        }

        try {
            const fileBuffer = fs.readFileSync(filePath);
            if (!hasUtf8Bom(fileBuffer)) {
                return;
            }

            const newText = document.getText();
            fs.writeFileSync(filePath, newText, { encoding: 'utf8' });

            const reopened = await vscode.workspace.openTextDocument(filePath);
            await vscode.window.showTextDocument(reopened, { preview: false });
            vscode.window.setStatusBarMessage('✅ UTF-8 BOM removed and saved', 3000);
        } catch (e) {
            console.error('Error checking BOM:', e);
        }
    });

    context.subscriptions.push(disposable);
}

export function deactivate() { }
