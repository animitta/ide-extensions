const assert = require('node:assert/strict');
const Module = require('node:module');
const path = require('node:path');
const { test } = require('node:test');
const { appendText, defaultTemplate, renderTemplate } = require('../dist/template');

function loadWithMocks(file, mocks) {
    const originalLoad = Module._load;
    const resolved = require.resolve(file);
    delete require.cache[resolved];
    Module._load = function (id, ...args) {
        return Object.hasOwn(mocks, id) ? mocks[id] : originalLoad.call(this, id, ...args);
    };
    try {
        return require(resolved);
    } finally {
        Module._load = originalLoad;
    }
}

test('renders default signature with a Git user and UTC across a year boundary', () => {
    const now = new Date('2027-01-01T00:30:45.123+08:00');
    assert.equal(renderTemplate(defaultTemplate, 'Alice', now),
        '// Checked by Alice 2026-12-31T16:30:45.123+00:00');
});

test('replaces all placeholders once, preserving literal user text and legacy timestamps', () => {
    const now = new Date('2026-11-02T08:30:45.007Z');
    assert.equal(renderTemplate('{user} {time} {{utc}} {user} $1\\path', '{time}$&', now),
        '{time}$& 2026-11-02T08:30:45.007+00:00 2026-11-02T08:30:45.007+00:00 {time}$& $1\\path');
});

test('wraps with one LF or CRLF and reuses an existing file newline', () => {
    for (const eol of ['\n', '\r\n']) {
        assert.equal(appendText('signature', false, eol), `${eol}signature${eol}`);
        assert.equal(appendText('signature', true, eol), `signature${eol}`);
        assert.equal(appendText('\n\nsignature\r\nsecond\n', true, eol), `signature${eol}second${eol}`);
    }
});

test('Git user lookup uses the file context and falls back on missing Git, errors or empty names', async () => {
    for (const [error, stdout, expected] of [
        [null, ' Alice \r\n', 'Alice'],
        [new Error('ENOENT'), '', 'Anonymous'],
        [new Error('exit 1'), '', 'Anonymous'],
        [new Error('timeout'), '', 'Anonymous'],
        [null, ' \n', 'Anonymous']
    ]) {
        const { getGitUser } = loadWithMocks('../dist/git', {
            'node:child_process': {
                execFile(command, args, options, callback) {
                    assert.equal(command, 'git');
                    assert.deepEqual(args, ['config', 'user.name']);
                    assert.equal(options.cwd, 'repo');
                    assert.equal(options.timeout, 3000);
                    callback(error, stdout);
                }
            }
        });
        assert.equal(await getGitUser('repo'), expected);
    }
});

function harness(content, { eol = '\n', template = defaultTemplate, accepted = true, noEditor = false } = {}) {
    const state = { information: [], warnings: [], subscriptions: [], edits: [] };
    const lines = content.split(/\r\n|\n/);
    const end = { line: lines.length - 1, character: lines.at(-1).length };
    const uri = { scheme: 'file', fsPath: path.join(process.cwd(), 'repo', 'file.txt') };
    const editor = {
        selections: [{ start: 0, end: 4 }, { start: 5, end: 5 }],
        document: {
            uri, lineCount: lines.length, eol: eol === '\r\n' ? 2 : 1,
            lineAt(index) {
                assert.equal(index, lines.length - 1);
                return { text: lines.at(-1), range: { end } };
            }
        },
        async edit(callback) {
            callback({ insert(position, text) { state.edits.push({ position, text }); } });
            return accepted;
        }
    };
    const vscode = {
        EndOfLine: { LF: 1, CRLF: 2 },
        commands: {
            registerCommand(id, callback) {
                state.command = id;
                state.run = callback;
                return { dispose() {} };
            }
        },
        window: {
            activeTextEditor: noEditor ? undefined : editor,
            showInformationMessage(message) { state.information.push(message); },
            showWarningMessage(message) { state.warnings.push(message); }
        },
        workspace: {
            getWorkspaceFolder() { return undefined; },
            getConfiguration(section, resource) {
                assert.equal(section, 'checker');
                assert.equal(resource, uri);
                return { get(key, fallback) {
                    assert.equal(key, 'template');
                    assert.equal(fallback, defaultTemplate);
                    return template;
                } };
            }
        }
    };
    loadWithMocks('../dist/extension', {
        vscode,
        './git': { async getGitUser(cwd) {
            assert.equal(cwd, path.dirname(uri.fsPath));
            return 'Alice';
        } }
    }).activate(state);
    return { state, end };
}

test('always inserts once at EOF regardless of selections, reusing existing line endings', async () => {
    for (const content of ['', 'body', 'body\n', 'body\n\n', 'body\r\n']) {
        const eol = content.includes('\r\n') ? '\r\n' : '\n';
        const { state, end } = harness(content, { eol });
        const before = Date.now();
        await state.run();
        const after = Date.now();
        assert.equal(state.edits.length, 1);
        assert.deepEqual(state.edits[0].position, end);
        const prefix = content.endsWith('\n') ? '' : eol;
        const text = state.edits[0].text;
        assert.ok(text.startsWith(`${prefix}// Checked by Alice `));
        assert.ok(text.endsWith(eol));
        const stamp = text.slice(`${prefix}// Checked by Alice `.length, -eol.length);
        assert.match(stamp, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}\+00:00$/);
        assert.ok(Date.parse(stamp) >= before && Date.parse(stamp) <= after);
    }
});

test('custom template remains literal and insertion rejection is reported', async () => {
    const { state } = harness('body', { template: '${1:value}\\path {user}', accepted: false });
    await state.run();
    assert.equal(state.edits[0].text, '\n${1:value}\\path Alice\n');
    assert.equal(state.warnings.length, 1);
});

test('registers a disposable and handles no active editor', async () => {
    const { state } = harness('', { noEditor: true });
    assert.equal(state.command, 'checker.insertTemplate');
    assert.equal(state.subscriptions.length, 1);
    await state.run();
    assert.equal(state.information.length, 1);
    assert.equal(state.edits.length, 0);
});
