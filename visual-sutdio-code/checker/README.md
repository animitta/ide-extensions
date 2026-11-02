# Checker

在当前文件末尾追加可配置的检查签名。默认模板为 `// Checked by {user} {time}`。

时间格式：`2026-11-02T08:30:45.123+00:00`，包含毫秒和明确的 UTC 时区偏移，不受电脑本地时区影响。

## 本地安装

1. 在 VS Code 扩展视图右上角的 `…` 菜单中选择 **从 VSIX 安装…**。
2. 选择 `checker-0.0.1.vsix`。

也可以执行：

```shell
code --install-extension checker-0.0.1.vsix
```

要求 VS Code 1.102.0 或更新版本。

## 使用

1. 打开需要添加检查签名的文本文件。
2. 按 `Ctrl+Shift+P`（macOS：`Cmd+Shift+P`）。
3. 运行 **Checker: 插入 UTC 字符串模板**。

也可以在编辑器中未选择文本时，依次按两次 `Ctrl+C`（组合快捷键 `Ctrl+C Ctrl+C`），直接执行插入操作。快捷键仅在可编辑的文本编辑器获得焦点且没有选区时生效。

默认追加内容示例：

```text
// Checked by Alice 2026-11-02T08:30:45.123+00:00
```

始终只在文件末尾追加一次，不受光标或选区影响。模板前后各加一个换行；文件已有结尾换行时复用，不额外添加前置换行。保留文件原有的空行，并沿用 LF/CRLF 格式。空文件也会插入前置换行。可以使用撤销恢复原文。

`{user}` 通过在当前文件所在目录执行 `git config user.name` 读取，遵循 Git 仓库级和全局配置。未保存文件使用所属工作区目录（无所属工作区时使用第一个本地工作区目录）。Git 不可用、读取失败、超时或用户名为空时使用 `Anonymous`。

## 自定义模板

在设置中搜索 `checker.template`，或者在用户/工作区 `settings.json` 中添加：

```json
{
  "checker.template": "// Checked by {user} {time}"
}
```

插入示例：

```text
// Checked by Alice 2026-11-02T08:30:45.123+00:00
```

`{user}`、`{time}` 可以出现多次，同次插入的时间一致。兼容旧的 `{{utc}}` 时间占位符。如果没有占位符，则插入静态模板。除占位符外，所有内容（包括 `$`、反斜杠和花括号）均按普通文本处理。模板内部支持换行，首尾换行统一为上述规则。

如需快捷键，可在 `keybindings.json` 中自行绑定命令：

```json
{
  "key": "ctrl+alt+u",
  "command": "checker.insertTemplate",
  "when": "editorTextFocus && !editorReadonly"
}
```

## 开发与打包

在本目录运行（Node.js 22 或更新版本）：

```shell
npm ci
npm test
npm run package
```

输出文件为本目录下的 `checker-0.0.1.vsix`。
