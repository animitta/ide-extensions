export const defaultTemplate = '// Checked by {user} {time}';

export function renderTemplate(template: string, user: string, now: Date = new Date()): string {
    const utc = now.toISOString().replace(/Z$/, '+00:00');
    return template.replace(/\{\{utc\}\}|\{time\}|\{user\}/g, token => token === '{user}' ? user : utc);
}

export function appendText(template: string, hasTrailingNewline: boolean, eol: string): string {
    const body = template.replace(/\r\n|\r|\n/g, '\n').replace(/^\n+|\n+$/g, '').replaceAll('\n', eol);
    return (hasTrailingNewline ? '' : eol) + body + eol;
}
