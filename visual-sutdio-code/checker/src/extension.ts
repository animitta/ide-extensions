import * as vscode from 'vscode';
import { dirname } from 'node:path';
import { appendText, defaultTemplate, renderTemplate } from './template';
import { getGitUser } from './git';

export function activate(context: vscode.ExtensionContext): void {
    context.subscriptions.push(vscode.commands.registerCommand('checker.insertTemplate', async () => {
        const editor = vscode.window.activeTextEditor;
        if (!editor) {
            vscode.window.showInformationMessage('请先打开一个文本编辑器，再插入 UTC 字符串模板。');
            return;
        }

        const document = editor.document;
        const folder = vscode.workspace.getWorkspaceFolder(document.uri);
        const cwd = document.uri.scheme === 'file' ? dirname(document.uri.fsPath)
            : folder?.uri.scheme === 'file' ? folder.uri.fsPath
            : vscode.workspace.workspaceFolders?.find(item => item.uri.scheme === 'file')?.uri.fsPath;
        const user = await getGitUser(cwd);
        if (document.isClosed) {
            return;
        }
        const template = vscode.workspace.getConfiguration('checker', document.uri)
            .get<string>('template', defaultTemplate);
        const lastLine = document.lineAt(document.lineCount - 1);
        const text = appendText(renderTemplate(template, user),
            document.lineCount > 1 && lastLine.text.length === 0,
            document.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n');
        const inserted = await editor.edit(builder => builder.insert(lastLine.range.end, text));
        if (!inserted) {
            vscode.window.showWarningMessage('无法插入 UTC 字符串模板，请确认当前文档可以编辑。');
        }
    }));
}
