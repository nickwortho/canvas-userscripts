// ==UserScript==
// @name         Export Canvas New Quiz
// @namespace    nw123
// @version      2026-09-22
// @description  Creates a native New Quizzes QTI export and converts it locally to Word document (.docx).
// @include      https://*.instructure.com/*
// @include      https://canvas.newcastle.edu.au/courses/*/assignments/*/edit?quiz_lti
// @include      https://canvas.newcastle.edu.au/courses/*/assignments/*/edit?quiz_lti=true
// @include      https://canvas.newcastle.edu.au/courses/*/assignments/*/edit?quiz_lti=true/
// @include      https://canvas.newcastle.edu.au/courses/*/assignments/*/build/*
// @include      https://canvas.newcastle.edu.au/courses/*/assignments/*/exports/*
// @require      https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js
// @require      https://cdn.jsdelivr.net/npm/docx@9.5.1/dist/index.iife.js
// @grant        GM_xmlhttpRequest
// @grant        GM_registerMenuCommand
// @connect      self
// @connect      instructure.com
// @connect      amazonaws.com
// @connect      cloudfront.net
// @connect      blob.core.windows.net
// @connect      canvas-user-content.com
// @run-at       document-idle
// ==/UserScript==


(function () {
    'use strict';

    // Fix JSZip/DOCX scheduling inside Tampermonkey
    window.setImmediate = (callback, ...args) =>
    window.setTimeout(() => callback(...args), 0);

    window.clearImmediate = id => window.clearTimeout(id);

    const CONFIG = {
        exportsLabel: /^Exports$/i,
        createLabel: /^\+?\s*Create Export$/i,
        downloadLabel: /^(?:Quiz|QTI) Export\b|^Download (?:Quiz |QTI )?Export\b/i,
        exportTimeoutMs: 180000,
        requestTimeoutMs: 60000,
        maxZipBytes: 100 * 1024 * 1024,
        maxXmlBytes: 12 * 1024 * 1024,
        maxTotalXmlBytes: 60 * 1024 * 1024,
        maxEntries: 5000
    };
    const VERSION = '1.0.0';
    const HOST_ID = 'nq-text-export-panel';
    const lname = node => (node && node.localName || '').toLowerCase();
    const children = node => Array.from(node && node.children || []);
    const descendants = (node, name) => Array.from(node ? node.getElementsByTagName('*') : [])
    .filter(el => lname(el) === name.toLowerCase());
    const first = (node, name) => descendants(node, name)[0] || null;
    const direct = (node, name) => children(node).find(el => lname(el) === name) || null;
    const content = node => (node && node.textContent || '').trim();
    const xml = node => new XMLSerializer().serializeToString(node);
    const clean = text => String(text || '').replace(/\r/g, '').replace(/\u00a0/g, ' ')
    .replace(/[ \t]+\n/g, '\n').replace(/\n[ \t]+/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
    const attrText = (el, excluded = []) => Array.from(el.attributes || [])
    .filter(a => !excluded.includes(a.name) && !a.name.startsWith('xmlns'))
    .map(a => a.name + '=' + JSON.stringify(a.value)).join(', ');
    const alphabet = n => {
        let s = '';
        for (n++; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
        return s;
    };
    const hasAncestor = (el, predicate, stop) => {
        for (let p = el.parentElement; p && p !== stop; p = p.parentElement) if (predicate(p)) return true;
        return false;
    };
    const isResponse = el => /^response_(lid|str|num|xy|grp)$/.test(lname(el));
    const sourceRef = value => /^data:/i.test(value || '') ? '[embedded data]' : (value || '');

    // A template is inert: quiz HTML is never inserted into the live page
    function htmlToText(html) {
        const template = document.createElement('template');
        template.innerHTML = html;
        function walk(node, pre = false) {
            if (node.nodeType === 3) return pre ? node.nodeValue : node.nodeValue.replace(/\s+/g, ' ');
            if (![1, 11].includes(node.nodeType)) return '';
            const tag = lname(node);
            if (/^(script|style|noscript)$/.test(tag)) return '';
            const body = () => Array.from(node.childNodes).map(n => walk(n, pre || tag === 'pre')).join('');
            if (tag === 'br') return '\n';
            if (tag === 'img') {
                const equation = node.getAttribute('data-equation-content');
                if (equation) return '[Equation: ' + equation + ']';
                return '[Image: ' + (node.getAttribute('alt') || 'no alternative text') +
                    (node.getAttribute('src') ? ' | ' + sourceRef(node.getAttribute('src')) : '') + ']';
            }
            if (/^(video|audio|iframe|object|embed)$/.test(tag)) {
                return '[' + tag + ': ' + sourceRef(node.getAttribute('src') || node.getAttribute('data') ||
                                                    node.querySelector('source')?.getAttribute('src')) + ']' + body();
            }
            if (tag === 'math') {
                return '[Math: ' + (node.getAttribute('alttext') ||
                                    content(node.querySelector('annotation[encoding*="tex"]')) || xml(node)) + ']';
            }
            if (tag === 'sup') return '^(' + body() + ')';
            if (tag === 'sub') return '_(' + body() + ')';
            if (tag === 'li') {
                const parent = node.parentElement;
                const marker = lname(parent) === 'ol' ? (children(parent).indexOf(node) +
                                                         Number(parent.getAttribute('start') || 1)) + '. ' : '- ';
                return '\n' + marker + body();
            }
            if (tag === 'a') {
                const label = body();
                const href = node.getAttribute('href') || '';
                return label + (href && href !== label && !/^(javascript:|#)/i.test(href) ?
                                ' (' + sourceRef(href) + ')' : '');
            }
            if (/^(td|th)$/.test(tag)) return body() + ' | ';
            if (/^(p|div|section|article|h[1-6]|ul|ol|table|tr|blockquote|pre)$/.test(tag)) {
                return '\n' + body() + '\n';
            }
            return body();
        }
        return clean(walk(template.content));
    }

    function materials(node, excludeResponses = false) {
        if (!node) return '';
        const leaves = Array.from(node.getElementsByTagName('*')).filter(el =>
                                                                         /^(mattext|matimage|mataudio|matvideo|matapplet|matref)$/.test(lname(el)) &&
                                                                         (!excludeResponses || !hasAncestor(el, p => isResponse(p) ||
                                                                                                            isResponse(node) && (/^render_/.test(lname(p)) || lname(p) === 'response_label'), node)));
        return clean(leaves.map(el => {
            if (lname(el) === 'mattext') {
                const raw = el.children.length ? Array.from(el.childNodes).map(xml).join('') : el.textContent;
                // Plain-text answers can contain <, > and regex syntax; never treat them as HTML
                return el.getAttribute('texttype')?.toLowerCase() === 'text/html' || el.children.length ?
                    htmlToText(raw) : clean(raw);
            }
            return '[' + lname(el).replace(/^mat/, '') + ': ' +
                sourceRef(el.getAttribute('uri') || el.getAttribute('linkrefid') || content(el)) + ']';
        }).join('\n'));
    }

    function metadata(item) {
        const values = new Map();
        // Keep repeated metadata rather than overwriting conflicting declarations.
        for (const field of descendants(item, 'qtimetadatafield')) {
            if (hasAncestor(field, p => lname(p) === 'item' && p !== item, item)) continue;
            const key = content(first(field, 'fieldlabel'));
            const value = content(first(field, 'fieldentry'));
            if (key) values.set(key, [...new Set([...(values.get(key) || []), value])]);
        }
        return values;
    }

    function renderItem(item, number, context, overrides = {}) {
        const meta = metadata(item);
        const type = (meta.get('question_type') || ['unspecified'])[0];
        const typeName = type.replace(/_question$/, '').replace(/_/g, ' ');
        const stimulus = type === 'text_only_question';
        const heading = stimulus ? 'STIMULUS / INFORMATION ' + number : 'QUESTION ' + number;
        const title = item.getAttribute('title') || 'Untitled';
        const points = overrides.points || (meta.get('points_possible') || [])[0];
        const lines = [heading + ': ' + title, 'Type: ' + typeName + (points ? ' | Points: ' + points : '')];
        const issues = [];
        const review = message => { if (!issues.includes(message)) issues.push(message); };
        const presentation = direct(item, 'presentation');
        const stem = materials(presentation, true);
        const parentStimulus = overrides.parent || item.getAttribute('parent_stimulus_item_ident');
        if (parentStimulus) lines.push('Related stimulus: ' + (context.titles.get(parentStimulus) || parentStimulus));
        if (item.getAttribute('instructions')) lines.push('Instructions: ' + htmlToText(item.getAttribute('instructions')));
        lines.push('', stem || '[Question text appears in the response fields below.]');
        if (!presentation) review('No presentation was supplied for this item.');

        const responses = new Map();
        const responseNodes = Array.from(presentation?.getElementsByTagName('*') || []).filter(isResponse);
        const blankType = /blank|dropdown/.test(type);
        for (let i = 0; i < responseNodes.length; i++) {
            const el = responseNodes[i];
            const id = el.getAttribute('ident') || 'response_' + (i + 1);
            const prompt = materials(el, true);
            const labels = descendants(el, 'response_label').filter(label =>
                                                                    !hasAncestor(label, p => isResponse(p) && p !== el, el));
            const options = labels.map((label, j) => ({
                id: label.getAttribute('ident'),
                letter: alphabet(j),
                text: materials(label) || content(label),
                node: label
            }));
            const name = blankType ? 'Blank ' + (i + 1) : (prompt || (responseNodes.length > 1 ? 'Response ' + (i + 1) : 'Answer'));
            const model = { id, name, options, node: el, ordered: el.getAttribute('rcardinality')?.toLowerCase() === 'ordered' };
            responses.set(id, model);
            // response_xy prompt/image live inside the response, not in the stem.
            if (lname(el) === 'response_xy') {
                lines.push('', materials(el));
            } else if (options.some(o => o.text)) {
                lines.push('', name + (model.ordered ? ' (arrange in order)' : '') + ':');
                if (blankType) lines.push('  Export response ID: ' + id);
                for (const option of options) {
                    const details = attrText(option.node, ['ident', 'position', 'lock']);
                    lines.push('  ' + option.letter + '. ' + (option.text || '[empty]') + (details ? ' [' + details + ']' : ''));
                }
            } else if (prompt) lines.push('', name);
        }
        if (blankType) {
            for (let i = 0; i < lines.length; i++) {
                for (const [id, response] of responses) lines[i] = lines[i].split('[' + id.replace(/^response_/, '') + ']').join('[' + response.name + ']');
            }
        }
        if (stimulus) {
            if (meta.has('source_url')) lines.push('', 'Source: ' + meta.get('source_url').join('; '));
        } else {
            lines.push('', 'ANSWER KEY');
            const calculated = first(item, 'calculated');
            const hotspot = first(presentation, 'render_hotspot');
            if (calculated) {
                lines.push('Formula question: answers depend on the generated variable values.');
                for (const formula of descendants(calculated, 'formula')) lines.push('Formula: ' + content(formula));
                const formulas = first(calculated, 'formulas');
                if (formulas && attrText(formulas)) lines.push('Formula settings: ' + attrText(formulas));
                for (const v of children(first(calculated, 'vars'))) {
                    lines.push('Variable ' + v.getAttribute('name') + ': ' + content(first(v, 'min')) +
                               ' to ' + content(first(v, 'max')) + (attrText(v, ['name']) ? ' [' + attrText(v, ['name']) + ']' : ''));
                }
                const tolerance = first(calculated, 'answer_tolerance');
                if (tolerance) lines.push('Tolerance: ' + content(tolerance) + (attrText(tolerance) ? ' [' + attrText(tolerance) + ']' : ''));
                for (const set of descendants(calculated, 'var_set')) {
                    const variables = children(set).filter(el => lname(el) === 'var')
                    .map(el => el.getAttribute('name') + '=' + content(el)).join(', ');
                    lines.push('  ' + variables + ' → ' + content(first(set, 'answer')));
                }
                // Preserve all extension parameters, including vendor fields not rendered above.
                lines.push('', 'Formula definition from QTI:', xml(calculated));
            } else if (hotspot) {
                lines.push('Correct region(s), using the coordinates stored in the export:');
                for (const label of descendants(hotspot, 'response_label')) {
                    lines.push('  ' + (label.getAttribute('rarea') || 'unspecified shape') + ': ' + (content(label) || '[missing coordinates]'));
                    const numbers = content(label).split(/[,\s]+/).map(Number);
                    if (!content(label) || numbers.length < 4 || numbers.length % 2 || numbers.some(n => !Number.isFinite(n))) {
                        review('Hot-spot coordinates are missing or invalid.');
                    }
                }
                review('Hot-spot image/regions require visual checking in Canvas; text cannot reproduce the image.');
            } else if (/^(essay|file_upload|file-upload)/.test(type)) {
                lines.push('Manual grading; no fixed correct answer is supplied by this question type.');
            } else {
                const rules = descendants(item, 'respcondition').filter(rule => descendants(rule, 'setvar').length ||
                                                                        /^(yes|true)$/i.test(rule.getAttribute('correctanswer') || ''));
                if (!rules.length) review('No interpretable answer/scoring rules were supplied.');
                function condition(node) {
                    if (!node) { review('Missing scoring condition.'); return '[missing condition]'; }
                    const tag = lname(node);
                    const operands = children(node).filter(el => lname(el) !== 'qticomment');
                    if (tag === 'not') return 'NOT (' + operands.map(condition).join(' AND ') + ')';
                    if (tag === 'and' || tag === 'or' || tag === 'conditionvar') {
                        const sameResponse = operands.length && operands.every(el => lname(el) === 'varequal' &&
                                                                               el.getAttribute('respident') === operands[0].getAttribute('respident'));
                        const response = sameResponse && responses.get(operands[0].getAttribute('respident'));
                        if (response?.ordered && tag !== 'or' && operands.every(el => !el.hasAttribute('index'))) {
                            return response.name + ', in order: ' + operands.map(el => answerValue(el, response)).join(' → ');
                        }
                        const values = operands.map(condition);
                        return values.length === 1 ? values[0] : '(' + values.join(tag === 'or' ? ' OR ' : ' AND ') + ')';
                    }
                    if (tag === 'other') return 'otherwise (QTI fallback condition)';
                    if (tag === 'unanswered') return (responses.get(node.getAttribute('respident'))?.name || 'Response') + ' is unanswered';
                    const ops = { varequal: '=', varlt: '<', varlte: '≤', vargt: '>', vargte: '≥', varsubset: 'is a subset of', varsubstring: 'contains' };
                    if (!ops[tag]) {
                        review('Unrecognised scoring operator: ' + tag + '.');
                        return '[QTI condition: ' + xml(node) + ']';
                    }
                    const response = responses.get(node.getAttribute('respident'));
                    if (!response) review('A scoring rule references an unknown response: ' + node.getAttribute('respident'));
                    const details = attrText(node, ['respident', 'index']);
                    const index = node.getAttribute('index');
                    const choice = response?.options.find(o => o.id === content(node));
                    const algorithm = choice?.node.getAttribute('scoring_algorithm');
                    const comparison = tag === 'varequal' ? ({
                        TextContainsAnswer: 'contains', TextCloseEnough: 'approximately matches',
                        TextRegex: 'matches regular expression', TextInChoices: 'is one of',
                        TextEquivalence: 'matches', Equivalence: '='
                    }[algorithm] || ops[tag]) : ops[tag];
                    return (response?.name || node.getAttribute('respident') || 'Answer') +
                        (index ? ' at position ' + index : '') + ' ' + comparison + ' ' + answerValue(node, response) +
                        (details ? ' [' + details + ']' : '');
                }
                function answerValue(node, response) {
                    const value = content(node);
                    const option = response?.options.find(o => o.id === value);
                    if (option) {
                        if (option.node.getAttribute('scoring_algorithm') === 'TextInChoices') {
                            return '{' + response.options.map(o => JSON.stringify(o.text)).join('; ') + '} [' +
                                attrText(option.node, ['ident', 'position', 'lock']) + ']';
                        }
                        const settings = attrText(option.node, ['ident', 'position', 'lock']);
                        return option.letter + '. ' + option.text + (settings ? ' [' + settings + ']' : '');
                    }
                    if (response && lname(response.node) === 'response_lid' && response.options.length) {
                        review('An answer identifier has no matching option: ' + value);
                    }
                    return JSON.stringify(value);
                }
                const declarations = descendants(first(item, 'resprocessing'), 'decvar');
                if (rules.length) {
//                    lines.push('The conditions below preserve the exported scoring rules, including partial credit.');
//                    for (const dec of declarations) lines.push('Score definition: ' + attrText(dec));
                    for (let i = 0; i < rules.length; i++) {
                        const rule = rules[i];
                        const cond = condition(first(rule, 'conditionvar'));
                        lines.push('  ' + (i + 1) + '. ' + cond);
//                        const award = descendants(rule, 'setvar').map(v =>
//                                                                      (v.getAttribute('action') || 'Set') + ' ' + (v.getAttribute('varname') || 'SCORE') + ' ' + content(v)).join('; ');
//                        const marked = /^(yes|true)$/i.test(rule.getAttribute('correctanswer') || '') ? ' [marked correct in export]' : '';
//                        lines.push('  ' + (i + 1) + '. ' + cond + ' → ' + (award || 'marked correct') + marked +
//                                   ' [continue=' + (rule.getAttribute('continue') || 'No (QTI default)') + ']');
                        const feedback = descendants(rule, 'displayfeedback').map(f => f.getAttribute('linkrefid')).filter(Boolean);
                        if (feedback.length) lines.push('     Feedback: ' + feedback.join(', '));
                    }
                }
            }
        }

        for (const feedback of children(item).filter(el => lname(el) === 'itemfeedback')) {
            const text = materials(feedback);
            if (text) lines.push('', 'Feedback [' + (feedback.getAttribute('ident') || 'general') + ']:', text);
        }
        const settings = Array.from(meta.entries()).filter(([key]) =>
                                                           !['question_type', 'points_possible', 'original_answer_ids', 'assessment_question_identifierref', 'source_url'].includes(key));
//        if (settings.length) lines.push('', 'Exported settings:', ...settings.map(([k, v]) => '  ' + k + ': ' + v.join('; ')));
        if ((meta.get('question_type') || []).length > 1) review('Conflicting question-type metadata; inspect the raw item.');
        if (first(item, 'itemproc_extension') && !first(item, 'calculated')) review('Additional vendor answer data is present; inspect the raw item.');
        if (!/^(unspecified|multiple_choice_question|choice_question|true_false_question|multiple_answers_question|multiple_answer_question|short_answer_question|fill_in_multiple_blanks_question|fill_in_the_blank_question|fill_in_blank_question|multiple_blanks_question|multiple_dropdowns_question|matching_question|categorization_question|ordering_question|numerical_question|calculated_question|formula_question|hot_spot_question|essay_question|file_upload_question|text_only_question)$/.test(type)) {
            review('Unrecognised question type: ' + type);
        }
        if (issues.length) {
            lines.push('', 'REVIEW REQUIRED:', ...issues.map(x => '- ' + x), '', 'Original QTI item:', xml(item));
            context.warnings.push(heading + ': ' + issues.join(' '));
        }
        return lines.join('\n');
    }

    function decodeXml(bytes) {
        const arr = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
        let encoding = 'utf-8';
        if (arr[0] === 0xff && arr[1] === 0xfe || arr[0] === 0x3c && arr[1] === 0) encoding = 'utf-16le';
        else if (arr[0] === 0xfe && arr[1] === 0xff || arr[0] === 0 && arr[1] === 0x3c) encoding = 'utf-16be';
        else {
            const header = new TextDecoder('ascii').decode(arr.slice(0, 200));
            encoding = header.match(/<\?xml[^>]*encoding\s*=\s*["']([^"']+)/i)?.[1] || encoding;
        }
        return new TextDecoder(encoding, { fatal: true }).decode(arr);
    }

    function parseXml(text, path) {
        if (/<!ENTITY/i.test(text)) throw new Error('Entity declarations are not supported: ' + path);
        const doc = new DOMParser().parseFromString(text, 'application/xml');
        if (first(doc, 'parsererror')) throw new Error('Malformed XML: ' + path);
        return { doc, path };
    }

    async function convertQti(buffer, filename = 'quiz.zip') {
        const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
        if (bytes.byteLength > CONFIG.maxZipBytes) throw new Error('Export exceeds the 100 MB limit.');
        const records = [];
        const warnings = [];
        if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
            if (typeof JSZip === 'undefined') throw new Error('JSZip did not load. Reinstall the script or allow its pinned @require dependency.');
            const zip = await JSZip.loadAsync(bytes);
            const files = Object.values(zip.files).filter(file => !file.dir);
            if (files.length > CONFIG.maxEntries) throw new Error('Too many files in the export.');
            let total = 0;
            for (const file of files.filter(f => /\.(xml|qti)$/i.test(f.name))) {
                const data = await file.async('uint8array');
                total += data.byteLength;
                if (data.byteLength > CONFIG.maxXmlBytes || total > CONFIG.maxTotalXmlBytes) throw new Error('XML content exceeds the configured size limits.');
                try { records.push(parseXml(decodeXml(data), file.name)); }
                catch (error) { warnings.push(error.message); }
            }
        } else {
            if (bytes.length > CONFIG.maxXmlBytes) throw new Error('XML exceeds the configured size limit.');
            records.push(parseXml(decodeXml(bytes), filename));
        }
        const qti = records.filter(r => ['questestinterop', 'assessment', 'objectbank', 'item'].includes(lname(r.doc.documentElement)));
        if (!qti.length) {
            if (records.some(r => /assessment(item|test)/.test(lname(r.doc.documentElement)))) {
                throw new Error('This is QTI 2/3. This script targets the QTI 1.2 export produced by Canvas New Quizzes.');
            }
            throw new Error('No readable QTI 1.2 found. The download may be a sign-in/error page. ' + warnings.join(' '));
        }
        const titles = new Map();
        const itemIndex = new Map();
        const banks = new Map();
        for (const record of qti) {
            for (const node of Array.from(record.doc.getElementsByTagName('*'))) {
                const id = node.getAttribute('ident');
                if (id && node.getAttribute('title')) titles.set(id, node.getAttribute('title'));
                if (lname(node) === 'item' && id) {
                    const list = itemIndex.get(id) || [];
                    list.push(node); itemIndex.set(id, list);
                }
                if (lname(node) === 'objectbank' && id) banks.set(id, node);
            }
        }
        const context = { warnings, titles };
        const assessments = qti.flatMap(r => descendants(r.doc, 'assessment').map(node => ({ node, path: r.path })));
        const title = assessments[0]?.node.getAttribute('title') || filename.replace(/\.(zip|xml|qti)$/i, '') || 'Quiz';
        let questionCount = 0, stimulusCount = 0, missingCount = 0;
        const sections = [];
        function visit(node, out, stack = new Set()) {
            const tag = lname(node);
            if (tag === 'item') {
                const stimulus = (metadata(node).get('question_type') || [])[0] === 'text_only_question';
                const n = stimulus ? ++stimulusCount : ++questionCount;
                out.push(renderItem(node, n, context), '');
                return;
            }
            if (tag === 'bankentry_item' || tag === 'itemref') {
                const id = node.getAttribute('item_ref') || node.getAttribute('linkrefid');
                const sourceBank = banks.get(node.getAttribute('sourcebank_ref'));
                const scoped = sourceBank && descendants(sourceBank, 'item').filter(el => el.getAttribute('ident') === id);
                const candidates = scoped?.length ? scoped : (itemIndex.get(id) || []);
                if (candidates.length === 1) {
                    out.push(renderItem(candidates[0], ++questionCount, context, {
                        points: node.getAttribute('points_possible'), parent: node.getAttribute('parent_stimulus_item_ident')
                    }), '');
                } else {
                    missingCount++;
                    const note = 'UNRESOLVED BANK ITEM: ' + id + ' — ' +
                          (candidates.length ? 'more than one definition found.' : 'question/answer content is absent from this QTI package.');
                    out.push(note, xml(node), '');
                    warnings.push(note);
                }
                return;
            }
            if (tag === 'section') {
                if (node.getAttribute('title')) out.push('ITEM BANK: ' + node.getAttribute('title'));//, '');
                const selection = first(direct(node, 'selection_ordering'), 'selection');
                if (selection) {
                    const count = content(first(selection, 'selection_number'));
                    out.push('RANDOM GROUP: choose ' + (count || 'an unspecified number of') + ' item(s). All possible items in bank follow.', '');
                    const bankId = content(first(selection, 'sourcebank_ref')) || content(first(selection, 'sourcebank_export_id'));
                    if (bankId && !descendants(node, 'item').length) {
                        const bank = banks.get(bankId);
                        if (bank && !stack.has(bankId)) {
                            const next = new Set(stack); next.add(bankId);
                            visit(bank, out, next);
                        } else {
                            const message = 'UNRESOLVED RANDOM BANK: ' + bankId + ' — candidate questions are not available in this package.';
                            out.push(message, ''); warnings.push(message); missingCount++;
                        }
                    }
                }
            }
            for (const child of children(node)) {
                if (['assessment', 'section', 'objectbank', 'item', 'bankentry_item', 'itemref'].includes(lname(child))) visit(child, out, stack);
            }
        }
        if (assessments.length) {
            for (const assessment of assessments) {
                const out = [];//['ASSESSMENT: ' + (assessment.node.getAttribute('title') || 'Untitled'), 'QTI source: ' + assessment.path, ''];
                const instructions = materials(direct(assessment.node, 'presentation_material'));
//                if (instructions) out.push('Instructions:', instructions, '');
                visit(assessment.node, out);
                sections.push(out.join('\n'));
            }
        } else {
            for (const record of qti) {
                const out = ['QTI source: ' + record.path, ''];
                visit(record.doc.documentElement, out);
                // questestinterop is a wrapper; visit() descends into all recognised children.
//                sections.push(out.join('\n'));
            }
        }
        if (!questionCount && !stimulusCount && !missingCount) throw new Error('The QTI package contains no question items.');
        const notice = warnings.length ? ['REVIEW REQUIRED (' + warnings.length + '):', ...warnings.map(w => '- ' + w), ''] : [];
        const text = [
            title, '='.repeat(Math.min(72, Math.max(8, title.length))),
//            'Questions/candidates: ' + questionCount + ' | Stimuli: ' + stimulusCount + ' | Unresolved references: ' + missingCount,
//            'Source: ' + filename, 'Converted: ' + new Date().toISOString(), 'Converter: ' + VERSION, '',
//            'Contains answer keys. Question/choice order is the exported order, not a randomised student attempt.',
//            'Score values below are QTI scoring units; use each question’s Points value for its quiz weighting.',
//            'Images/media remain text references. Relative paths refer to the original QTI ZIP; retain it if needed.',
//            'Only content actually included in the export can be reproduced.', '',
//            ...notice,
            ...sections
        ].join('\n');
        return { text, title, questionCount, stimulusCount, missingCount, warnings };
    }

    // Native UI adapter. No guessed/private Canvas API endpoints are used.
    const isVisible = el => Boolean(el && el.getClientRects().length &&
                                    getComputedStyle(el).visibility !== 'hidden' && !el.closest('[hidden],[aria-hidden="true"]'));
    const label = el => clean(el.getAttribute('aria-label') || el.textContent).replace(/\s+/g, ' ');
    const controls = () => Array.from(document.querySelectorAll('button,a,[role="button"],[role="tab"]'))
    .filter(el => !el.closest('#' + HOST_ID) && isVisible(el));
    const findControl = regex => controls().find(el => regex.test(label(el)));
    const usable = el => el && !el.disabled && el.getAttribute('aria-disabled') !== 'true';
    function isEditor() {
        const labels = controls().map(label);
        // Require the New Quizzes navigation rather than matching Canvas course exports.
        return labels.some(x => /^Build$/i.test(x)) && labels.some(x => /^Moderate$/i.test(x)) &&
            labels.some(x => CONFIG.exportsLabel.test(x));
    }
    const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
    function checkAbort(signal) { if (signal?.aborted) throw new DOMException('Cancelled', 'AbortError'); }
    async function until(test, ms, signal, message) {
        const deadline = Date.now() + ms;
        while (Date.now() < deadline) {
            checkAbort(signal);
            const found = test();
            if (found) return found;
            await pause(300);
        }
        throw new Error(message);
    }
    function downloadLinks() {
        return Array.from(document.querySelectorAll('a[href]')).filter(el =>
                                                                       isVisible(el) && CONFIG.downloadLabel.test(label(el)) && /^https?:|^blob:/.test(el.href));
    }
    function exportKey(link) {
        // Ignore rotating signatures in old links. New exports have new file paths.
        const url = new URL(link.href, location.href);
        for (const name of Array.from(url.searchParams.keys())) {
            if (/^(x-amz-|x-goog-|signature$|expires$|awsaccesskeyid$|token$|sig$|se$|st$|sv$|sp$|sr$)/i.test(name)) url.searchParams.delete(name);
        }
        return url.origin + url.pathname + url.search;
    }
    function editorIdentity() {
        // Exports navigation usually changes only the tab segment.
        const query = new URLSearchParams(location.search);
        query.delete('tab');
        return (location.pathname + '?' + query + location.hash)
            .replace(/\/(build|edit|exports|settings|reports|moderate)(?=\/|$|\?)/g, '/TAB');
    }
    async function nativeExport(signal, status) {
        checkAbort(signal);
        if (!isEditor()) throw new Error('Open the New Quizzes Build/Exports page as an instructor first.');
        if (!findControl(CONFIG.createLabel)) {
            status('Opening the Exports tab…');
            const tab = findControl(CONFIG.exportsLabel);
            if (!usable(tab)) throw new Error('The native Exports tab is unavailable.');
            tab.click();
        }
        await until(() => usable(findControl(CONFIG.createLabel)), 15000, signal,
                    'Create Export was not found. Open Exports manually, then retry.');
        status('Checking the export list…');
        // Allow the initial asynchronous export list to settle before recording its baseline.
        let prior = '', stableSince = Date.now();
        await until(() => {
            const signature = downloadLinks().map(exportKey).sort().join('|');
            const loading = Array.from(document.querySelectorAll('[role="progressbar"],[aria-busy="true"]')).some(isVisible);
            if (signature !== prior || loading) { prior = signature; stableSince = Date.now(); }
            return !loading && Date.now() - stableSince > 2000;
        }, 20000, signal, 'The export list is still loading. Wait for Canvas, then retry.');
        const before = new Set(downloadLinks().map(exportKey));
        const button = findControl(CONFIG.createLabel);
        if (!usable(button)) throw new Error('Create Export is disabled. Wait for any existing export to finish.');
        const pending = Array.from(document.querySelectorAll('tr,[role="row"],li')).some(el =>
                                                                                         isVisible(el) && /(?:quiz|qti) export/i.test(el.textContent) &&
                                                                                         /(?:in progress|processing|pending|queued|exporting)/i.test(el.textContent));
        if (pending) throw new Error('Another export is still in progress. Wait for it to finish, then retry.');
        const identity = editorIdentity();
        status('Creating a new QTI export. Keep this quiz open…');
        checkAbort(signal);
        button.click(); // Exactly once. Never retry a mutation automatically.
        const link = await until(() => {
            if (!isEditor() || editorIdentity() !== identity) throw new Error('The quiz page changed. Export monitoring stopped.');
            const candidates = downloadLinks().filter(el => !before.has(exportKey(el)));
            if (candidates.length > 1) throw new Error('More than one new export appeared. Use Convert downloaded QTI to choose the correct file.');
            return candidates[0];
        }, CONFIG.exportTimeoutMs, signal,
                                 'No new downloadable export appeared within 3 minutes. If Canvas finished, download its QTI ZIP and use Convert downloaded QTI.');
        return { url: link.href, name: link.getAttribute('download') || 'canvas-new-quiz.zip' };
    }

    async function getExport(url, signal) {
        checkAbort(signal);
        const target = new URL(url, location.href);
        if (!['https:', 'blob:'].includes(target.protocol)) throw new Error('Unexpected export URL scheme.');
        if (target.origin === location.origin || target.protocol === 'blob:') {
            const controller = new AbortController();
            const abort = () => controller.abort();
            signal?.addEventListener('abort', abort, { once: true });
            const timer = setTimeout(abort, CONFIG.requestTimeoutMs);
            try {
                const response = await fetch(target.href, { credentials: 'same-origin', signal: controller.signal });
                if (!response.ok) throw new Error('Export download failed (HTTP ' + response.status + ').');
                return await response.arrayBuffer();
            } catch (error) {
                checkAbort(signal);
                if (target.protocol === 'blob:') throw error;
                // Signed redirects can require the userscript's cross-origin downloader.
            } finally {
                clearTimeout(timer);
                signal?.removeEventListener('abort', abort);
            }
        }
        if (typeof GM_xmlhttpRequest !== 'function') throw new Error('Cross-origin download is unavailable. Download the QTI ZIP manually and convert it with the file button.');
        return await new Promise((resolve, reject) => {
            let request;
            const abort = () => { request?.abort(); reject(new DOMException('Cancelled', 'AbortError')); };
            const done = callback => value => { signal?.removeEventListener('abort', abort); callback(value); };
            signal?.addEventListener('abort', abort, { once: true });
            request = GM_xmlhttpRequest({
                method: 'GET', url: target.href, responseType: 'arraybuffer',
                anonymous: target.origin !== location.origin,
                timeout: CONFIG.requestTimeoutMs,
                onload: done(response => {
                    if (response.status >= 200 && response.status < 300) resolve(response.response);
                    else reject(new Error('Export download failed (HTTP ' + response.status + '). Download the QTI manually and use the file button.'));
                }),
                onerror: done(() => reject(new Error('Download blocked or unavailable. Allow the export download host in Tampermonkey, or convert the manually downloaded QTI.'))),
                ontimeout: done(() => reject(new Error('Export download timed out.'))),
                onabort: done(() => reject(new DOMException('Cancelled', 'AbortError')))
            });
        });
    }

    function saveFile(blob, filename, target) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url; link.download = filename; link.textContent = 'Download ' + filename;
        // Retain a real clickable link in case automatic iframe downloads are blocked.
        target.append(link, document.createElement('br'));
        link.click();
        return url;
    }

    const safeName = name => (name || 'quiz').replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/[. ]+$/g, '').slice(0, 130) || 'quiz';
    let panel = null;
    let active = null;
    let mountedAt = '';
    const objectUrls = [];

    function mount(force = false) {
        if (panel || !document.body || (!force && !isEditor())) return;
        const host = document.createElement('div');
        host.id = HOST_ID;
        const root = host.attachShadow({ mode: 'open' });
        root.innerHTML = '<style>' +
            ':host{position:fixed;right:18px;bottom:18px;z-index:2147483000;font:14px/1.4 system-ui,sans-serif;color:#172b3a}' +
            'section{background:#fff;border:1px solid #b8c5cf;border-radius:9px;box-shadow:0 3px 20px #0002;padding:14px;width:290px}' +
            'strong{display:block;margin-bottom:9px}button{display:block;width:100%;margin:7px 0;padding:9px;border:1px solid #718798;border-radius:5px;background:#f2f6f8;cursor:pointer;font:inherit}' +
            'button.primary{background:#075b91;color:#fff}button:disabled{opacity:.55;cursor:wait}p{margin:9px 0;font-size:12px;overflow-wrap:anywhere}' +
            'label{display:block;font-size:12px}a{font-size:12px;color:#075b91}details summary{cursor:pointer;font-size:12px}[hidden]{display:none!important}' +
            '</style><section aria-label="Quiz text export"><strong>Quiz → text + answer key</strong>' +
            '<button class="primary" id="export">Export quiz as DOCX</button>' +
            '<button id="file">Convert downloaded QTI…</button><input id="input" type="file" accept=".zip,.xml,.qti" hidden>' +
            '<label><input id="retain" type="checkbox"> Also download the original QTI</label>' +
            '<button id="cancel" hidden>Stop waiting</button>' +
            '<p id="status" role="status" aria-live="polite">Uses Canvas’s native quiz export. Answers are included.</p>' +
            '<div id="downloads"></div><details><summary>Help</summary><p>Open Build → Exports if the export control is not detected. If automatic export fails, download the QTI ZIP from Canvas and choose it above. Conversion stays in your browser.</p></details></section>';
        document.body.append(host);
        panel = { host, root };
        mountedAt = location.href;
        const $ = id => root.getElementById(id);
        const status = message => { $('status').textContent = message; };

        async function run(file) {
            if (active) return;
            const controller = new AbortController();
            active = controller;
            $('export').disabled = $('file').disabled = true; $('cancel').hidden = false;
            $('downloads').replaceChildren();
            objectUrls.splice(0).forEach(url => URL.revokeObjectURL(url));
            try {
                let data, filename;
                if (file) {
                    if (file.size > CONFIG.maxZipBytes) throw new Error('File exceeds the 100 MB limit.');
                    status('Reading QTI…'); data = await file.arrayBuffer(); filename = file.name;
                } else {
                    const exported = await nativeExport(controller.signal, status);
                    status('Downloading the new QTI…');
                    data = await getExport(exported.url, controller.signal); filename = exported.name;
                }
                checkAbort(controller.signal); status('Converting questions and answer keys…');
                const result = await convertQti(data, filename);
                checkAbort(controller.signal);
                const name = safeName(result.title);
                // Word docx formatting
                status('Creating Word document…');

                const wordDocument = new docx.Document({
                    sections: [{
                        children: result.text.split(/\r?\n/).map(line =>
                                                                 new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: line,
                                    font: 'Calibri',
                                    size: 22 // 11 pt; DOCX uses half-point units
                                })
                            ],
                            spacing: { after: 0 }
                        })
                                                                )
                    }]
                });

                const wordBlob = await docx.Packer.toBlob(wordDocument);
                checkAbort(controller.signal);

                objectUrls.push(saveFile(
                    wordBlob,
                    name + '-questions-and-answers.docx',
                    $('downloads')
                ));
                //
                if ($('retain').checked) objectUrls.push(saveFile(new Blob([data], {
                    type: 'application/octet-stream'
                }), name + (new Uint8Array(data)[0] === 0x50 ? '.zip' : '.xml'), $('downloads')));
                status('Converted ' + result.questionCount + ' question(s)/candidates. ' +
                       (result.warnings.length ? result.warnings.length + ' review note(s): see the top of the text file.' : 'Download ready.') +
                       ' If no download started, use the link below.');
            } catch (error) {
                status(error.name === 'AbortError' ? 'Stopped. A Canvas export already started may still finish.' : error.message);
            } finally {
                active = null; $('export').disabled = $('file').disabled = false; $('cancel').hidden = true;
            }
        }
        $('export').addEventListener('click', () => run(null));
        $('file').addEventListener('click', () => $('input').click());
        $('input').addEventListener('change', event => {
            const file = event.target.files[0]; event.target.value = ''; if (file) run(file);
        });
        $('cancel').addEventListener('click', () => active?.abort());
    }

    // CommonJS is used only by the local test harness; userscript managers do not define it.
    if (typeof module !== 'undefined' && module.exports) {
        module.exports = { convertQti, htmlToText, renderItem, nativeExport, exportKey, decodeXml, CONFIG };
        return;
    }
    if (typeof GM_registerMenuCommand === 'function') GM_registerMenuCommand('Open QTI → text converter', () => mount(true));
    let queued = false;
    function refresh() {
        queued = false;
        if (panel && !active && location.href !== mountedAt && !isEditor()) {
            panel.host.remove(); panel = null;
            objectUrls.splice(0).forEach(url => URL.revokeObjectURL(url));
        }
        mount();
    }
    const observer = new MutationObserver(() => {
        if (!queued) { queued = true; setTimeout(refresh, 250); }
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });
    mount();
})();
