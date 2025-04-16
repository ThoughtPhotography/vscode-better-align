import * as vscode from 'vscode';

enum TokenType {
    Invalid = 'Invalid',
    Word = 'Word',
    Assignment = 'Assignment', // = += -= *= /= %= ~= |= ^= .= :=
    Arrow = 'Arrow', // =>
    Block = 'Block', // {} [] ()
    PartialBlock = 'PartialBlock', // { [ (
    EndOfBlock = 'EndOfBlock', // } ] )
    String = 'String',
    PartialString = 'PartialString',
    Comment = 'Comment',
    Whitespace = 'Whitespace',
    Colon = 'Colon',
    Comma = 'Comma',
    CommaAsWord = 'CommaAsWord',
    Insertion = 'Insertion',
}

interface Token {
    type: TokenType;
    text: string;
}

export interface LineInfo {
    line: vscode.TextLine;
    sgfntTokenType: TokenType;
    sgfntTokens: TokenType[];
    tokens: Token[];
}

export interface LineRange {
    anchor: number;
    infos: LineInfo[];
}

const REG_WS = /\s/;
const BRACKET_PAIR: any = {
    '{': '}',
    '[': ']',
    '(': ')',
};

function whitespace(count: number) {
    return new Array(count + 1).join(' ');
}

export class Formatter {
    /* Align:
     *   operators = += -= *= /= :
     *   trailling comment
     *   preceding comma
     * Ignore anything inside a quote, comment, or block
     */
    public process(editor: vscode.TextEditor, ignoreGaps: boolean = false): void {
        this.editor = editor;

        // Get line ranges
        const ranges = this.getLineRanges(editor, ignoreGaps);

        // Format
        let formatted: string[][] = [];
        for (let range of ranges) {
            formatted.push(this.format(range, ignoreGaps));
        }

        // Apply
        editor.edit((editBuilder) => {
            for (let i = 0; i < ranges.length; ++i) {
                var infos = ranges[i].infos;
                var lastline = infos[infos.length - 1].line;
                var location = new vscode.Range(infos[0].line.lineNumber, 0, lastline.lineNumber, lastline.text.length);
                const eol = editor.document.eol === vscode.EndOfLine.LF ? '\n' : '\r\n';
                const replaced = formatted[i].join(eol);
                if (editor.document.getText(location) === replaced) {
                    continue;
                }
                editBuilder.replace(location, replaced);
            }
        });
    }

    protected editor: vscode.TextEditor;

    protected getLineRanges(editor: vscode.TextEditor, ignoreGaps: boolean): LineRange[] {
        var ranges: LineRange[] = [];
        editor.selections.forEach((sel) => {
            const indentBase = this.getConfig().get('indentBase', 'firstline') as string;
            const importantIndent: boolean = indentBase === 'dontchange';

            let res: LineRange;
            if (sel.isSingleLine) {
                // If this selection is single line. Look up and down to search for the similar neighbour
                res = ignoreGaps
                    ? this.narrowIgnoringGaps(0, editor.document.lineCount - 1, sel.active.line, importantIndent)
                    : this.narrow(0, editor.document.lineCount - 1, sel.active.line, importantIndent);
                if (res.infos.length > 0) { // Ensure narrow returned something
                    ranges.push(res);
                }
            } else {
                // Otherwise, narrow down the range where to align
                let start = sel.start.line;
                let end = sel.end.line;

                if (ignoreGaps) {
                    // Treat entire selection as one block, tokenizing all lines
                    let infos: LineInfo[] = [];
                    let commonTokenTypes: TokenType[] | null = null;
                    for (let i = start; i <= end; i++) {
                        const tokenInfo = this.tokenize(i);
                        // TODO: Add check here to skip blank/comment lines if needed for type checking?
                        // For now, just collect all lines.
                        infos.push(tokenInfo);
                        if (tokenInfo.sgfntTokens.length > 0) {
                            if (commonTokenTypes === null) {
                                commonTokenTypes = tokenInfo.sgfntTokens;
                            } else {
                                commonTokenTypes = this.arrayAnd(commonTokenTypes, tokenInfo.sgfntTokens);
                            }
                        }
                    }
                    // Basic filtering: Only proceed if there are lines with common significant tokens
                    if (infos.length > 0 && commonTokenTypes && commonTokenTypes.length > 0) {
                        // Assign the most common significant token type (e.g., Assignment > Colon)
                        let sgt = TokenType.Assignment;
                        if (commonTokenTypes.indexOf(TokenType.Assignment) === -1) {
                             sgt = commonTokenTypes[0]; // Or choose based on priority
                        }
                        infos.forEach(info => info.sgfntTokenType = sgt);
                        ranges.push({ anchor: sel.anchor.line, infos });
                    }
                } else {
                    // Original behavior: Iterate using narrow
                    while (true) {
                        res = this.narrow(start, end, start, importantIndent);
                        if (!res.infos[0]) {
                            break; // Avoid infinite loop if narrow returns empty
                        }
                        let lastLine = res.infos[res.infos.length - 1];

                        if (lastLine.line.lineNumber > end) {
                            break;
                        }

                        if (res.infos[0] && res.infos[0].sgfntTokenType !== TokenType.Invalid) {
                            ranges.push(res);
                        }

                        if (lastLine.line.lineNumber === end) {
                            break;
                        }

                        start = lastLine.line.lineNumber + 1;
                    }
                }
            }
        });
        return ranges;
    }

    protected getConfig() {
        let defaultConfig = vscode.workspace.getConfiguration('betterAlign');
        let langConfig: any = null;

        try {
            langConfig = vscode.workspace.getConfiguration().get(`[${this.editor.document.languageId}]`) as any;
        } catch (e) {}

        return {
            get: function (key: any, defaultValue?: any): any {
                if (langConfig) {
                    var key1 = 'betterAlign.' + key;
                    if (langConfig.hasOwnProperty(key1)) {
                        return langConfig[key1];
                    }
                }

                return defaultConfig.get(key, defaultValue);
            },
        };
    }

    protected tokenize(line: number): LineInfo {
        let textline = this.editor.document.lineAt(line);
        let text = textline.text;
        let pos = 0;
        let lt: LineInfo = {
            line: textline,
            sgfntTokenType: TokenType.Invalid,
            sgfntTokens: [],
            tokens: [],
        };

        let lastTokenType = TokenType.Invalid;
        let tokenStartPos = -1;

        while (pos < text.length) {
            let char = text.charAt(pos);
            let next = text.charAt(pos + 1);
            let third = text.charAt(pos + 2);

            let currTokenType: TokenType;

            let nextSeek = 1;

            // Tokens order are important
            if (char.match(REG_WS)) {
                currTokenType = TokenType.Whitespace;
            } else if (char === '"' || char === "'" || char === '`') {
                currTokenType = TokenType.String;
            } else if (char === '{' || char === '(' || char === '[') {
                currTokenType = TokenType.Block;
            } else if (char === '}' || char === ')' || char === ']') {
                currTokenType = TokenType.EndOfBlock;
            } else if (
                char === '/' &&
                ((next === '/' && (pos > 0 ? text.charAt(pos - 1) : '') !== ':') || // only `//` but not `://`
                    next === '*')
            ) {
                currTokenType = TokenType.Comment;
            } else if (char === ',') {
                if (lt.tokens.length === 0 || (lt.tokens.length === 1 && lt.tokens[0].type === TokenType.Whitespace)) {
                    currTokenType = TokenType.CommaAsWord; // Comma-first style
                } else {
                    currTokenType = TokenType.Comma;
                }
            } else if (char === '=' && next === '>') {
                currTokenType = TokenType.Arrow;
                nextSeek = 2;
            } else if (
                // Currently we support only known operators,
                // formatters will not work for unknown operators, we should find a way to support all operators.
                // Math operators
                (char === '+' ||
                    char === '-' ||
                    char === '*' ||
                    char === '/' ||
                    char === '%' || // FIXME: Find a way to work with the `**` operator
                    // Bitwise operators
                    char === '~' ||
                    char === '|' ||
                    char === '^' || // FIXME: Find a way to work with the `<<` and `>>` bitwise operators
                    // Other operators
                    char === '.' ||
                    char === ':' ||
                    char === '!' ||
                    char === '&' ||
                    char === '=') &&
                next === '='
            ) {
                currTokenType = TokenType.Assignment;
                nextSeek = third === '=' ? 3 : 2;
            } else if (char === '=' && next !== '=') {
                currTokenType = TokenType.Assignment;
            } else if (char === ':' && next === ':') {
                currTokenType = TokenType.Word;
                nextSeek = 2;
            } else if (char === ':' && next !== ':' || (char === '?' && next === ':')) {
                currTokenType = TokenType.Colon;
            } else {
                currTokenType = TokenType.Word;
            }

            if (currTokenType !== lastTokenType) {
                if (tokenStartPos !== -1) {
                    lt.tokens.push({
                        type: lastTokenType,
                        text: textline.text.substr(tokenStartPos, pos - tokenStartPos),
                    });
                }

                lastTokenType = currTokenType;
                tokenStartPos = pos;

                if (
                    lastTokenType === TokenType.Assignment ||
                    lastTokenType === TokenType.Colon ||
                    lastTokenType === TokenType.Arrow ||
                    lastTokenType === TokenType.Comment
                ) {
                    if (lt.sgfntTokens.indexOf(lastTokenType) === -1) {
                        lt.sgfntTokens.push(lastTokenType);
                    }
                }
            }

            // Skip to end of string
            if (currTokenType === TokenType.String) {
                ++pos;
                while (pos < text.length) {
                    let quote = text.charAt(pos);
                    if (quote === char && text.charAt(pos - 1) !== '\\') {
                        break;
                    }
                    ++pos;
                }
                if (pos >= text.length) {
                    lastTokenType = TokenType.PartialString;
                }
            }

            // Skip to end of block
            if (currTokenType === TokenType.Block) {
                ++pos;
                let bracketCount = 1;
                while (pos < text.length) {
                    let bracket = text.charAt(pos);
                    if (bracket === char) {
                        ++bracketCount;
                    } else if (bracket === BRACKET_PAIR[char] && text.charAt(pos - 1) !== '\\') {
                        if (bracketCount === 1) {
                            break;
                        } else {
                            --bracketCount;
                        }
                    }
                    ++pos;
                }
                if (pos >= text.length) {
                    lastTokenType = TokenType.PartialBlock;
                }
                // -1 then + nextSeek so keep pos not change in next loop
                // or we will lost symbols like "] } )"
            }

            if (char === '/') {
                // Skip to end if we encounter single line comment
                if (next === '/') {
                    pos = text.length;
                } else if (next === '*') {
                    ++pos;
                    while (pos < text.length) {
                        if (text.charAt(pos) === '*' && text.charAt(pos + 1) === '/') {
                            ++pos;
                            currTokenType = TokenType.Word;
                            break;
                        }
                        ++pos;
                    }
                }
            }

            pos += nextSeek;
        }

        if (tokenStartPos !== -1) {
            lt.tokens.push({
                type: lastTokenType,
                text: textline.text.substr(tokenStartPos, pos - tokenStartPos),
            });
        }

        return lt;
    }

    protected hasPartialToken(info: LineInfo): boolean {
        for (let j = info.tokens.length - 1; j >= 0; --j) {
            let lastT = info.tokens[j];
            if (
                lastT.type === TokenType.PartialBlock ||
                lastT.type === TokenType.EndOfBlock ||
                lastT.type === TokenType.PartialString
            ) {
                return true;
            }
        }
        return false;
    }

    protected hasSameIndent(info1: LineInfo, info2: LineInfo): boolean {
        var t1 = info1.tokens[0];
        var t2 = info2.tokens[0];

        if (t1.type === TokenType.Whitespace) {
            if (t1.text === t2.text) {
                return true;
            }
        } else if (t2.type !== TokenType.Whitespace) {
            return true;
        }

        return false;
    }

    protected arrayAnd(array1: TokenType[], array2: TokenType[]): TokenType[] {
        var res: TokenType[] = [];
        var map: any = {};
        for (var i = 0; i < array1.length; ++i) {
            map[array1[i]] = true;
        }
        for (var i = 0; i < array2.length; ++i) {
            if (map[array2[i]]) {
                res.push(array2[i]);
            }
        }
        return res;
    }

    /*
     * Determine which blocks of code needs to be align.
     * 1. Empty lines is the boundary of a block.
     * 2. If user selects something, blocks are always within selection,
     *    but not necessarily is the selection.
     * 3. Bracket / Brace usually means boundary.
     * 4. Unsimilar line is boundary.
     */
    protected narrow(start: number, end: number, anchor: number, importantIndent: boolean): LineRange {
        let anchorToken = this.tokenize(anchor);
        if (this.isBlankOrComment(anchorToken)) {
             return { anchor, infos: [] }; // Don't align single blank/comment lines
        }
        let range = { anchor, infos: [anchorToken] };

        let tokenTypes = anchorToken.sgfntTokens;

        if (anchorToken.sgfntTokens.length === 0) {
            return range;
        }

        if (this.hasPartialToken(anchorToken)) {
            return range;
        }

        let i = anchor - 1;
        while (i >= start) {
            let token = this.tokenize(i);

            if (this.isBlankOrComment(token)) {
                break; // Original behavior stops at blank lines
            }

            if (this.hasPartialToken(token)) {
                break;
            }

            let tt = this.arrayAnd(tokenTypes, token.sgfntTokens);
            if (tt.length === 0) {
                break;
            }
            tokenTypes = tt;

            if (importantIndent && !this.hasSameIndent(anchorToken, token)) {
                break;
            }

            range.infos.unshift(token);
            --i;
        }

        i = anchor + 1;
        while (i <= end) {
            let token = this.tokenize(i);

            if (this.isBlankOrComment(token)) {
                break; // Original behavior stops at blank lines
            }

            let tt = this.arrayAnd(tokenTypes, token.sgfntTokens);
            if (tt.length === 0) {
                break;
            }
            tokenTypes = tt;

            if (importantIndent && !this.hasSameIndent(anchorToken, token)) {
                break;
            }

            if (this.hasPartialToken(token)) {
                range.infos.push(token);
                break;
            }

            range.infos.push(token);
            ++i;
        }

        let sgt;
        if (tokenTypes.indexOf(TokenType.Assignment) >= 0) {
            sgt = TokenType.Assignment;
        } else {
            sgt = tokenTypes[0];
        }
        for (let info of range.infos) {
            info.sgfntTokenType = sgt;
        }

        return range;
    }

    protected isBlankOrComment(info: LineInfo): boolean {
        if (info.line.isEmptyOrWhitespace) {
            return true;
        }
        // Check if tokens contain only whitespace and/or comments
        for (const token of info.tokens) {
            if (token.type !== TokenType.Whitespace && token.type !== TokenType.Comment) {
                return false;
            }
        }
        // If we only found whitespace or comments (or nothing), it's effectively blank/comment
        return true;
    }

    // New function to narrow range while ignoring gaps (blank/comment lines)
    protected narrowIgnoringGaps(start: number, end: number, anchor: number, importantIndent: boolean): LineRange {
        let anchorToken = this.tokenize(anchor);
        let initialAnchor = anchor; // Store original anchor for potential return if no valid line found
        // If anchor itself is blank/comment, find the nearest non-blank/comment line to use as anchor
        let searchDir = -1;
        while (this.isBlankOrComment(anchorToken)) {
            anchor += searchDir;
            if (anchor < start || anchor > end) {
                return { anchor: initialAnchor, infos: [] }; // No alignable lines found, return empty with original anchor
            }
            anchorToken = this.tokenize(anchor);
            // If we searched down and failed, try searching up
            if (searchDir === -1 && anchor < start) {
                searchDir = 1;
                anchor = anchorToken.line.lineNumber + 1; // Start searching upwards from original anchor + 1
            }
        }

        let range = { anchor, infos: [anchorToken] };
        let tokenTypes = anchorToken.sgfntTokens;

        if (anchorToken.sgfntTokens.length === 0 || this.hasPartialToken(anchorToken)) {
            return range; // Return just the valid anchor if it has no tokens or is partial
        }

        // Search upwards, skipping gaps
        let i = anchor - 1;
        while (i >= start) {
            let token = this.tokenize(i);

            if (this.isBlankOrComment(token)) {
                // Add blank/comment line to preserve it, but don't check its type or indent
                range.infos.unshift(token);
                i--;
                continue;
            }

            if (this.hasPartialToken(token)) {
                break; // Stop if we hit a partial token
            }

            let tt = this.arrayAnd(tokenTypes, token.sgfntTokens);
            if (tt.length === 0) {
                break; // Stop if token types don't match
            }
            tokenTypes = tt;

            if (importantIndent && !this.hasSameIndent(anchorToken, token)) {
                break; // Stop if indent doesn't match (when `indentBase` is 'dontchange')
            }

            range.infos.unshift(token);
            --i;
        }

        // Search downwards, skipping gaps
        i = anchor + 1;
        while (i <= end) {
            let token = this.tokenize(i);

            if (this.isBlankOrComment(token)) {
                // Add blank/comment line to preserve it, but don't check its type or indent
                range.infos.push(token);
                i++;
                continue;
            }

            let tt = this.arrayAnd(tokenTypes, token.sgfntTokens);
            if (tt.length === 0) {
                break; // Stop if token types don't match
            }
            tokenTypes = tt;

            if (importantIndent && !this.hasSameIndent(anchorToken, token)) {
                break; // Stop if indent doesn't match (when `indentBase` is 'dontchange')
            }

            // Don't include a partial token at the end of the downwards search
            // because it might belong to the next block
            if (this.hasPartialToken(token)) {
                break;
            }

            range.infos.push(token);
            ++i;
        }

        // Assign the determined significant token type to all non-gap lines
        let sgt;
        if (tokenTypes.indexOf(TokenType.Assignment) >= 0) {
            sgt = TokenType.Assignment;
        } else {
            sgt = tokenTypes.length > 0 ? tokenTypes[0] : TokenType.Invalid; // Default if somehow no common types found
        }
        for (let info of range.infos) {
            if (!this.isBlankOrComment(info)) {
                info.sgfntTokenType = sgt;
            }
        }

        return range;
    }

    protected format(range: LineRange, ignoreGaps: boolean = false): string[] {
        // 0. Handle empty range
        if (!range.infos.length) {
            return [];
        }

        const config = this.getConfig();
        const indentBaseSetting = config.get('indentBase', 'firstline') as string;

        // --- Store Original Lines & Identify Lines to Format ---
        const originalLines: { [key: number]: LineInfo } = {};
        range.infos.forEach(info => originalLines[info.line.lineNumber] = info);

        let linesToFormat: LineInfo[] = [];
        if (ignoreGaps) {
            linesToFormat = range.infos.filter(info => !this.isBlankOrComment(info));
        } else {
            linesToFormat = range.infos; // Keep all lines if not ignoring gaps
        }

        // If, after filtering, there are no lines left to format, return original lines
        if (linesToFormat.length === 0) {
            return range.infos.map(info => info.line.text);
        }

        // --- Determine Indentation ---
        let indentation = '';
        let firstNonSpaceCharIndex = 0;
        let whiteSpaceType = ' '; // Default to space
        const firstLineToFormat = linesToFormat[0];

        if (firstLineToFormat.tokens.length > 0 && firstLineToFormat.tokens[0].type === TokenType.Whitespace) {
            whiteSpaceType = firstLineToFormat.tokens[0].text[0] ?? ' ';
        }

        if (ignoreGaps || indentBaseSetting === 'firstline') {
            // Use first *alignable* line's indent
            firstNonSpaceCharIndex = firstLineToFormat.line.text.search(/\S/);
            if (firstNonSpaceCharIndex === -1) {firstNonSpaceCharIndex = 0; {// Handle lines with only whitespace somehow? Should be caught by isBlankOrComment
                indentation = whiteSpaceType.repeat(firstNonSpaceCharIndex);
            }
        } else if (indentBaseSetting === 'activeline') {
            // Find anchor line's indent among the alignable lines
            let anchorLine = linesToFormat.find(info => info.line.lineNumber === range.anchor) ?? firstLineToFormat;
            firstNonSpaceCharIndex = anchorLine.line.text.search(/\S/);
            if (firstNonSpaceCharIndex === -1) {firstNonSpaceCharIndex = 0;} {
                indentation = whiteSpaceType.repeat(firstNonSpaceCharIndex);
            }
        } else { // 'dontchange' - applies only if not ignoring gaps, otherwise handled by narrow logic
            // Indentation will be handled line-by-line based on original
            // We still need the whitespace type from the first line
            indentation = ''; // Placeholder, will be set per line later
        }

        // --- Prepare Lines for Formatting ---
        // Remove indent and trailing whitespace from lines *to be formatted*
        linesToFormat.forEach(info => {
            if (info.tokens[0]?.type === TokenType.Whitespace) {
                info.tokens.shift();
            }
            if (info.tokens.length > 0 && info.tokens[info.tokens.length - 1].type === TokenType.Whitespace) {
                info.tokens.pop();
            }
        });

        // --- Calculate First Word Length (for specific alignment cases) ---
        let firstWordLength = 0;
        // Only calculate if not ignoring gaps (original behavior)
        if (!ignoreGaps) {
            for (let info of linesToFormat) {
                let count = 0;
                for (let token of info.tokens) {
                    if (token.type === info.sgfntTokenType) {
                        count = -count;
                        break;
                    }
                    if (token.type === TokenType.Block) {
                        continue;
                    }
                    if (token.type !== TokenType.Whitespace) {
                        ++count;
                    }
                }
                if (count < -1) {
                    firstWordLength = Math.max(firstWordLength, info.tokens[0]?.text.length ?? 0);
                }
            }
        }

        // --- Remove Whitespace Around Operators ---
        for (let info of linesToFormat) {
            let i = 1;
            while (i < info.tokens.length) {
                if (info.tokens[i].type === info.sgfntTokenType || info.tokens[i].type === TokenType.Comma) {
                    if (info.tokens[i - 1]?.type === TokenType.Whitespace) {
                        info.tokens.splice(i - 1, 1);
                        --i;
                    }
                    if (info.tokens[i + 1]?.type === TokenType.Whitespace) {
                        info.tokens.splice(i + 1, 1);
                    }
                }
                ++i;
            }
        }

        // --- Align Logic ---
        const configOP = config.get('operatorPadding') as string;
        const configWS = config.get('surroundSpace');
        const stt = TokenType[linesToFormat[0].sgfntTokenType]?.toLowerCase() ?? 'assignment'; // Default if invalid
        const configDef: any = {
            colon: [0, 1],
            assignment: [1, 1],
            comment: 2,
            arrow: [1, 1],
        };
        const configSTT = configWS[stt] || configDef[stt];
        const configComment = configWS['comment'] || configDef['comment'];

        const formatSize = linesToFormat.length;
        let length = new Array<number>(formatSize).fill(0);
        let column = new Array<number>(formatSize).fill(0);
        let formattedContent = new Array<string>(formatSize).fill(''); // Store content after indentation

        let exceed = 0;
        let hasTrallingComment = false;
        let resultSize = 0; // Tracks max length before operator

        while (exceed < formatSize) {
            let operatorSize = 0;

            // First pass: scan to next operator for each line being formatted
            for (let l = 0; l < formatSize; ++l) {
                let i = column[l];
                let info = linesToFormat[l];
                let tokenSize = info.tokens.length;

                if (i === -1) {
                    continue; // Already finished this line
                }

                let end = tokenSize;
                let lineContent = formattedContent[l];

                // Check for trailing comment
                if (tokenSize > 0 && info.tokens[tokenSize - 1].type === TokenType.Comment) {
                    hasTrallingComment = true;
                    end = tokenSize - 1;
                    if (tokenSize > 1 && info.tokens[tokenSize - 2].type === TokenType.Whitespace) {
                        end = tokenSize - 2;
                    }
                }

                for (; i < end; ++i) {
                    let token = info.tokens[i];
                    if (token.type === info.sgfntTokenType || (token.type === TokenType.Comma && i !== 0)) {
                        operatorSize = Math.max(operatorSize, token.text.length);
                        break;
                    } else {
                        lineContent += token.text;
                    }
                }

                formattedContent[l] = lineContent;
                if (i < end) {
                    // Only update resultSize if we stopped at an operator within the main content
                    resultSize = Math.max(resultSize, lineContent.length);
                }

                if (i === end) {
                    // Reached end (or trailing comment start)
                    ++exceed;
                    column[l] = -1;
                    // Store remaining tokens (like the trailing comment) for later
                    info.tokens = info.tokens.slice(end);
                } else {
                    // Stopped at an operator
                    column[l] = i;
                }
            }

            // Second pass: align operators
            if (exceed >= formatSize) {
                break; // Exit if all lines finished in first pass
            }

            for (let l = 0; l < formatSize; ++l) {
                let i = column[l];
                if (i === -1) {
                    continue;
                }

                let info = linesToFormat[l];
                let lineContent = formattedContent[l];

                let op = info.tokens[i].text;
                if (op.length < operatorSize) {
                    if (configOP === 'right') {
                        op = whitespace(operatorSize - op.length) + op;
                    } else {
                        op = op + whitespace(operatorSize - op.length);
                    }
                }

                let padding = '';
                if (resultSize > lineContent.length) {
                    padding = whitespace(resultSize - lineContent.length);
                }

                if (info.tokens[i].type === TokenType.Comma) {
                    lineContent += op;
                    if (i < info.tokens.length - 1) {
                        lineContent += padding + ' '; // Ensure one space after comma
                    }
                } else {
                    // Apply surround space settings
                    if (configSTT[0] < 0) {
                        // Stick left
                        if (configSTT[1] < 0) {
                            // Stick both - complex alignment (original logic)
                            let z = lineContent.length - 1;
                            while (z >= 0) {
                                let ch = lineContent.charAt(z);
                                if (ch.match(REG_WS)) {
                                    break;
                                }
                                --z;
                            }
                            lineContent = lineContent.substring(0, z + 1) + padding + lineContent.substring(z + 1) + op;
                        } else {
                            lineContent = lineContent + op; // Stick left only
                            if (i < info.tokens.length - 1) {
                                 lineContent += padding; // Add padding before next token
                            }
                        }
                    } else {
                        // Space before operator
                        lineContent = lineContent + padding + whitespace(configSTT[0]) + op;
                    }
                    if (configSTT[1] > 0) {
                        // Space after operator
                        lineContent += whitespace(configSTT[1]);
                    }
                }

                formattedContent[l] = lineContent;
                column[l] = i + 1; // Move past the operator for the next iteration
                // We processed an operator, so this line isn't finished yet.
                // Find the index of this line in the exceed calculation and potentially remove it
                // This logic might be complex, simpler to just recalculate exceed based on column[l] === -1
            }
            // Recalculate exceed count after processing operators
            exceed = column.filter(c => c === -1).length;
        }

        // --- Align Trailing Comments ---
        if (hasTrallingComment && configComment >= 0) {
            resultSize = 0;
            for (let l = 0; l < formatSize; ++l) {
                resultSize = Math.max(resultSize, formattedContent[l].length);
            }
            for (let l = 0; l < formatSize; ++l) {
                let info = linesToFormat[l];
                if (info.tokens.length > 0 && info.tokens[0].type === TokenType.Comment) {
                    let lineContent = formattedContent[l];
                    formattedContent[l] = lineContent + whitespace(resultSize - lineContent.length + configComment) + info.tokens.join('');
                }
            }
        } else {
            // Append remaining tokens (comments) without alignment if configComment < 0 or no comments found
            for (let l = 0; l < formatSize; ++l) {
                let info = linesToFormat[l];
                if (info.tokens.length > 0) {
                    formattedContent[l] += info.tokens.join('');
                }
            }
        }

        // --- Reconstruct Final Output ---
        const finalResult: string[] = [];
        const formattedLinesMap = new Map<number, string>();
        linesToFormat.forEach((info, index) => {
            // Apply final indentation
             let finalIndentation = indentation; // Use calculated base indentation
            if (!ignoreGaps && indentBaseSetting === 'dontchange') {
                // For 'dontchange' without ignoring gaps, use original indent
                const originalInfo = originalLines[info.line.lineNumber];
                finalIndentation = originalInfo.line.text.substring(0, originalInfo.line.text.search(/\S|$/));
            }
            // Ensure finalIndentation is a string
            if (typeof finalIndentation !== 'string') {
                finalIndentation = ''; // Default to empty string if undefined/null
            }
            formattedLinesMap.set(info.line.lineNumber, finalIndentation + formattedContent[index]);
        });

        // Iterate through the original lines in order
        range.infos.forEach(originalInfo => {
            if (formattedLinesMap.has(originalInfo.line.lineNumber)) {
                finalResult.push(formattedLinesMap.get(originalInfo.line.lineNumber)!);
            } else {
                // This was a blank/comment line, push its original text
                finalResult.push(originalInfo.line.text);
            }
        });

        return finalResult;
    }
}
}
