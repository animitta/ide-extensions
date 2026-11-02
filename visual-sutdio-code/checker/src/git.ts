import { execFile } from 'node:child_process';

export function getGitUser(cwd?: string): Promise<string> {
    return new Promise(resolve => {
        execFile('git', ['config', 'user.name'], {
            cwd,
            encoding: 'utf8',
            timeout: 3000,
            windowsHide: true
        }, (error, stdout) => {
            const user = stdout.trim().replace(/[\r\n]+/g, ' ');
            resolve(error || !user ? 'Anonymous' : user);
        });
    });
}
