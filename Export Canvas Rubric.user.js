// ==UserScript==
// @name         Export Canvas Rubric
// @namespace    nw123
// @version      2026-10-22
// @description  Creates a rubric export locally formatting to Word doc (.docx)
// @include      https://canvas.newcastle.edu.au/courses/*/rubrics*
// @grant        none
// @run-at       document-idle
// ==/UserScript==


(function () {
  'use strict';

  const PREFIX = 'canvas-rubric-docx';
  const ITEM = `[data-${PREFIX}-item]`;
  const TRIGGER =
    '[data-testid^="rubric-options-"][data-testid$="-button"]';
  const MENU_MARKER = '[data-testid="edit-rubric-button"]';
  const MIME =
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  const courseId = () =>
    location.pathname.match(/^\/courses\/(\d+)\/rubrics\/?$/)?.[1];

  if (!courseId() || document.getElementById(`${PREFIX}-style`)) return;

  let active = null;
  let frame = 0;

  const owners = new WeakMap();
  const jobs = new Set();

  const style = document.createElement('style');
  style.id = `${PREFIX}-style`;
  style.textContent = `
    .${PREFIX}-row {
      list-style: none;
      margin: 0;
      padding: 0;
    }

    .${PREFIX}-item {
      appearance: none;
      display: block;
      box-sizing: border-box;
      width: 100%;
      border: 0;
      border-radius: 0;
      margin: 0;
      padding: 0.5rem 0.75rem;
      background: transparent;
      color: var(--ic-brand-font-color-dark, #2d3b45);
      font: inherit;
      line-height: 1.5;
      text-align: start;
      cursor: pointer;
    }

    .${PREFIX}-item:hover,
    .${PREFIX}-item:focus {
      background: var(--ic-brand-primary, #0374b5);
      color: #fff;
      outline: 2px solid transparent;
      outline-offset: -2px;
    }

    .${PREFIX}-item[aria-disabled="true"] {
      opacity: 0.55;
      cursor: wait;
    }

    #${PREFIX}-status {
      position: fixed;
      bottom: 24px;
      right: 24px;
      z-index: 2147483647;
      max-width: min(440px, 85vw);
      padding: 12px 16px;
      border: 1px solid #c7cdd1;
      border-radius: 4px;
      background: #fff;
      color: #2d3b45;
      box-shadow: 0 2px 12px #0003;
      font: 14px/1.5 sans-serif;
    }
  `;
  document.head.appendChild(style);

  function visible(el) {
    return (
      !!el &&
      el.isConnected &&
      el.getClientRects().length > 0 &&
      getComputedStyle(el).visibility !== 'hidden'
    );
  }

  function rememberTrigger(target) {
    const button =
      target instanceof Element ? target.closest(TRIGGER) : null;

    const id = button
      ?.getAttribute('data-testid')
      ?.match(/^rubric-options-(\d+)-button$/)?.[1];

    const course = courseId();

    if (!id || !course) return false;

    const row = button.closest(
      '[data-testid^="rubric-row-"],tr,[role="row"]'
    );

    active = {
      id,
      course,
      button,
      hidePoints:
        row
          ?.querySelector('[data-testid^="rubric-points-"]')
          ?.textContent.trim() === '--',
    };

    schedule();
    return true;
  }

  function schedule() {
    if (!frame) {
      frame = requestAnimationFrame(() => {
        frame = 0;
        installItem();
      });
    }
  }

  function findMenu() {
    if (
      !active?.button.isConnected ||
      active.course !== courseId()
    ) {
      return null;
    }

    const menus = [...document.querySelectorAll(MENU_MARKER)]
      .map(el => el.closest('[role="menu"]'))
      .filter(visible);

    const unique = [...new Set(menus)];

    const ids = [
      'aria-controls',
      'aria-owns',
      'aria-describedby',
    ]
      .flatMap(name =>
        (active.button.getAttribute(name) || '').split(/\s+/)
      )
      .filter(Boolean);

    for (const id of ids) {
      const panel = document.getElementById(id);

      const match = unique.find(
        menu => panel === menu || panel?.contains(menu)
      );

      if (match) return match;
    }

    const labelled =
      active.button.id &&
      unique.find(menu =>
        (menu.getAttribute('aria-labelledby') || '')
          .split(/\s+/)
          .includes(active.button.id)
      );

    if (labelled) return labelled;

    // Avoid guessing if more than one rubric menu is visible.
    return unique.length === 1 ? unique[0] : null;
  }

  function installItem() {
    const menu = findMenu();
    if (!menu) return;

    const existing = menu.querySelector(ITEM);

    if (existing) {
      const owner = owners.get(existing);

      if (
        owner?.id === active.id &&
        owner.course === active.course
      ) {
        return;
      }

      existing.closest(`.${PREFIX}-row`)?.remove();
    }

    const row = document.createElement(
      /^(UL|OL)$/.test(menu.tagName) ? 'li' : 'div'
    );

    row.className = `${PREFIX}-row`;
    row.setAttribute('role', 'none');

    const button = document.createElement('button');

    button.type = 'button';
    button.className = `${PREFIX}-item`;
    button.setAttribute(`data-${PREFIX}-item`, '');
    button.setAttribute('role', 'menuitem');
    button.setAttribute('aria-label', 'Export rubric to Word');
    button.tabIndex = -1;
    button.textContent = 'Export';

    owners.set(button, { ...active, menu });

    row.appendChild(button);
    menu.appendChild(row);
  }

  function stop(event) {
    event.preventDefault();
    event.stopImmediatePropagation();
  }

  function closeMenu(owner) {
    if (visible(owner.menu) && owner.button.isConnected) {
      owner.button.click();
    }

    owner.button.focus({ preventScroll: true });
  }

  document.addEventListener(
    'click',
    event => {
      const item =
        event.target instanceof Element
          ? event.target.closest(ITEM)
          : null;

      if (item) {
        stop(event);

        const owner = owners.get(item);

        if (owner && owner.course === courseId()) {
          void exportRubric(owner);
        }

        return;
      }

      rememberTrigger(event.target);
    },
    true
  );

  document.addEventListener(
    'keydown',
    event => {
      if (
        ['Enter', ' ', 'ArrowDown', 'ArrowUp'].includes(event.key)
      ) {
        rememberTrigger(event.target);
      }

      if (!(event.target instanceof Element)) return;

      const menu = event.target.closest('[role="menu"]');
      const item = menu?.querySelector(ITEM);

      if (!item) return;

      const owner = owners.get(item);
      if (!owner) return;

      const onExport =
        event.target === item || item.contains(event.target);

      if (onExport && ['Enter', ' '].includes(event.key)) {
        stop(event);

        if (!event.repeat) item.click();
      } else if (
        ['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)
      ) {
        const items = [
          ...menu.querySelectorAll(
            '[role="menuitem"], ' +
            '[role="menuitemcheckbox"], ' +
            '[role="menuitemradio"]'
          ),
        ].filter(
          el =>
            visible(el) &&
            el.getAttribute('aria-disabled') !== 'true' &&
            !el.disabled
        );

        if (!items.length) return;

        stop(event);

        let index = items.findIndex(
          el => el === event.target || el.contains(event.target)
        );

        if (event.key === 'Home') {
          index = 0;
        } else if (event.key === 'End') {
          index = items.length - 1;
        } else {
          index =
            (index +
              (event.key === 'ArrowDown' ? 1 : -1) +
              items.length) %
            items.length;
        }

        items[index].focus();
      } else if (onExport && event.key === 'Escape') {
        stop(event);
        closeMenu(owner);
      }
    },
    true
  );

  new MutationObserver(schedule).observe(document.body, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: [
      'aria-expanded',
      'aria-controls',
      'hidden',
    ],
  });

  window.addEventListener('popstate', () => {
    active = null;
  });

  rememberTrigger(document.activeElement);

  let toastTimer;

  function status(
    message,
    error = false,
    persistent = false
  ) {
    let el = document.getElementById(`${PREFIX}-status`);

    if (!el) {
      el = document.createElement('div');
      el.id = `${PREFIX}-status`;
      el.setAttribute('aria-live', 'polite');
      document.body.appendChild(el);
    }

    el.setAttribute('role', error ? 'alert' : 'status');
    el.style.borderColor = error ? '#b52828' : '#c7cdd1';
    el.textContent = message;

    clearTimeout(toastTimer);

    if (!persistent) {
      toastTimer = setTimeout(
        () => el.remove(),
        error ? 15000 : 6000
      );
    }
  }

  async function loadRubric(owner) {
    const controller = new AbortController();
    const timeout = setTimeout(
      () => controller.abort(),
      30000
    );

    try {
      const url = new URL(
        `/api/v1/courses/${owner.course}/rubrics/${owner.id}`,
        location.origin
      );

      // Include display settings, without student assessments.
      url.searchParams.append(
        'include[]',
        'course_associations'
      );

      const response = await fetch(url, {
        method: 'GET',
        credentials: 'same-origin',
        signal: controller.signal,
        headers: {
          Accept: 'application/json',
        },
      });

      if (!response.ok) {
        const hints = {
          401:
            'Your Canvas session has expired. Sign in again and retry.',
          403:
            'Canvas denied access. Your role needs permission to manage rubrics.',
          404:
            'Canvas could not find this rubric in the course. Refresh the Rubrics page.',
          429:
            'Canvas is rate limiting requests. Wait a moment and retry.',
        };

        throw new Error(
          hints[response.status] ||
          `Canvas returned HTTP ${response.status}. Please retry.`
        );
      }

      if (
        !(response.headers.get('content-type') || '')
          .includes('json')
      ) {
        throw new Error(
          'Canvas returned a sign-in page instead of rubric data. Sign in again and retry.'
        );
      }

      const payload = await response.json();
      const rubric = payload.rubric || payload;

      if (
        String(rubric.id) !== owner.id ||
        !Array.isArray(rubric.data)
      ) {
        throw new Error(
          'Canvas returned an unexpected rubric response. No document was exported.'
        );
      }

      const association = (
        rubric.associations || []
      ).find(
        a =>
          a.association_type === 'Course' &&
          String(a.association_id) === owner.course &&
          a.purpose === 'bookmark'
      );

      return {
        ...rubric,
        hide_points:
          owner.hidePoints ||
          truth(rubric.hide_points) ||
          truth(association?.hide_points),
        hide_score_total:
          truth(rubric.hide_score_total) ||
          truth(association?.hide_score_total),
      };
    } catch (error) {
      if (error.name === 'AbortError') {
        throw new Error(
          'Canvas did not respond within 30 seconds. Please retry.'
        );
      }

      throw error;
    } finally {
      clearTimeout(timeout);
    }
  }

  async function exportRubric(owner) {
    const key = `${owner.course}:${owner.id}`;

    if (jobs.has(key)) {
      status('This rubric is already being exported.');
      return;
    }

    jobs.add(key);
    closeMenu(owner);

    status('Exporting rubric to Word…', false, true);

    try {
      const rubric = await loadRubric(owner);
      const blob = makeDocx(rubric);
      const url = URL.createObjectURL(blob);

      const link = document.createElement('a');
      link.href = url;
      link.download = filename(rubric.title, owner.id);
      link.style.display = 'none';

      document.body.appendChild(link);
      link.click();
      link.remove();

      setTimeout(() => URL.revokeObjectURL(url), 60000);

      status(
        'Word document ready. Check your browser downloads.'
      );
    } catch (error) {
      console.error('[Canvas Rubric Export]', error);

      status(
        `Export failed: ${
          error.message || 'Unknown error. Please retry.'
        }`,
        true
      );
    } finally {
      jobs.delete(key);
    }
  }

  const truth = value =>
    value === true ||
    value === 1 ||
    value === 'true' ||
    value === '1';

  const number = value =>
    value === '' ||
    value == null ||
    !Number.isFinite(Number(value))
      ? null
      : Number(value);

  const pointText = value =>
    number(value) == null ? '—' : String(number(value));

  function filename(title, id) {
    let safe = plain(title)
      .replace(/[<>:"/\\|?*\u0000-\u001f]/g, ' ')
      .replace(/\s+/g, ' ')
      .replace(/\.docx$/i, '')
      .trim()
      .slice(0, 140)
      .replace(/[. ]+$/, '');

    if (!safe) safe = `Rubric ${id}`;

    if (
      /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(
        safe
      )
    ) {
      safe = `Rubric ${safe}`;
    }

    return `${safe}.docx`;
  }

  function plain(value) {
    const raw = String(value ?? '').replace(/\r\n?/g, '\n');

    const hasMarkup =
      /<\/?(?:p|div|br|span|strong|em|b|i|u|s|a|ul|ol|li|h[1-6]|table|tr|td|th|img|sup|sub|blockquote|pre|code|script|style|iframe|object|svg|math)\b|&(?:#\d+|#x[\da-f]+|[a-z]+);/i;

    if (!hasMarkup.test(raw)) return raw.trim();

    // Template contents are inert and never enter the live page.
    const template = document.createElement('template');
    template.innerHTML = raw;

    const blocks = new Set([
      'P', 'DIV', 'UL', 'OL', 'LI',
      'H1', 'H2', 'H3', 'H4', 'H5', 'H6',
      'TR', 'BLOCKQUOTE', 'PRE',
    ]);

    function walk(node) {
      if (node.nodeType === 3) return node.textContent;

      if (node.nodeType !== 1 && node.nodeType !== 11) {
        return '';
      }

      const tag = node.nodeName;

      if (
        [
          'SCRIPT', 'STYLE', 'IFRAME',
          'OBJECT', 'EMBED', 'NOSCRIPT',
        ].includes(tag)
      ) {
        return '';
      }

      if (tag === 'BR') return '\n';

      if (tag === 'IMG') {
        const label =
          node.getAttribute('alt') ||
          node.getAttribute('data-equation-content');

        return label ? `[Image: ${label}]` : '[Image]';
      }

      const body = [...node.childNodes].map(walk).join('');
      const prefix = tag === 'LI' ? '• ' : '';

      return (
        (blocks.has(tag) && tag !== 'LI' ? '\n' : '') +
        prefix +
        body +
        (
          blocks.has(tag)
            ? '\n'
            : ['TD', 'TH'].includes(tag)
              ? '\t'
              : ''
        )
      );
    }

    return walk(template.content)
      .replace(/\u00a0/g, ' ')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n[ \t]+/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  }

  function xml(value) {
    return String(value ?? '')
      .replace(
        /[^\u0009\u000a\u000d\u0020-\ud7ff\ue000-\ufffd\u{10000}-\u{10ffff}]/gu,
        ''
      )
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function paragraph(
    value,
    {
      style: styleName = '',
      bold = false,
      centre = false,
      keep = false,
    } = {}
  ) {
    const properties =
      (
        styleName
          ? `<w:pStyle w:val="${xml(styleName)}"/>`
          : ''
      ) +
      (keep ? '<w:keepNext/>' : '') +
      (centre ? '<w:jc w:val="center"/>' : '');

    const runs = String(value ?? '')
      .split('\n')
      .map(
        (line, index) =>
          `<w:r>` +
          (bold ? '<w:rPr><w:b/></w:rPr>' : '') +
          (index ? '<w:br/>' : '') +
          `<w:t xml:space="preserve">${xml(line)}</w:t>` +
          `</w:r>`
      )
      .join('');

    return (
      `<w:p><w:pPr>${properties}</w:pPr>` +
      (runs || '<w:r><w:t/></w:r>') +
      `</w:p>`
    );
  }

  function table(headers, rows, widths) {
    function row(values, isHeader) {
      return (
        `<w:tr>` +
        (
          isHeader
            ? '<w:trPr><w:tblHeader/></w:trPr>'
            : ''
        ) +
        values
          .map(
            (text, i) =>
              `<w:tc><w:tcPr>` +
              `<w:tcW w:w="${widths[i]}" w:type="dxa"/>` +
              (
                isHeader
                  ? '<w:shd w:val="clear" w:fill="E8EFF5"/>'
                  : ''
              ) +
              `<w:vAlign w:val="center"/></w:tcPr>` +
              paragraph(text, {
                bold: isHeader,
                centre: headers[i] === 'Points',
                style: 'TableText',
                keep: isHeader,
              }) +
              `</w:tc>`
          )
          .join('') +
        `</w:tr>`
      );
    }

    const borders = [
      'top', 'left', 'bottom', 'right',
      'insideH', 'insideV',
    ]
      .map(
        side =>
          `<w:${side} w:val="single" w:sz="4" w:color="D9D9D9"/>`
      )
      .join('');

    return (
      `<w:tbl><w:tblPr>` +
      `<w:tblW w:w="9360" w:type="dxa"/>` +
      `<w:tblLayout w:type="fixed"/>` +
      `<w:tblBorders>${borders}</w:tblBorders>` +
      `<w:tblCellMar>` +
      `<w:top w:w="100" w:type="dxa"/>` +
      `<w:left w:w="120" w:type="dxa"/>` +
      `<w:bottom w:w="100" w:type="dxa"/>` +
      `<w:right w:w="120" w:type="dxa"/>` +
      `</w:tblCellMar></w:tblPr>` +
      `<w:tblGrid>` +
      widths
        .map(width => `<w:gridCol w:w="${width}"/>`)
        .join('') +
      `</w:tblGrid>` +
      row(headers, true) +
      rows.map(values => row(values, false)).join('') +
      `</w:tbl>`
    );
  }

  function ratingRows(criterion, order, hidePoints) {
    const ratings = [...(criterion.ratings || [])].sort(
      (a, b) =>
        (number(b.points) ?? 0) -
        (number(a.points) ?? 0)
    );

    const rows = ratings.map((rating, index) => {
      const high = number(rating.points);

      const low =
        index + 1 < ratings.length
          ? number(ratings[index + 1].points)
          : 0;

      const range =
        truth(criterion.criterion_use_range) &&
        high != null &&
        low != null &&
        high !== 0 &&
        high !== low;

      const points = range
        ? `${pointText(high)} to >${pointText(
            Math.round(low * 100) / 100
          )}`
        : pointText(rating.points);

      const values = [
        plain(rating.description),
        plain(rating.long_description),
      ];

      return hidePoints ? values : [points, ...values];
    });

    return order === 'ascending' ? rows.reverse() : rows;
  }

  function makeDocx(rubric) {
    const hidePoints = truth(rubric.hide_points);
    const freeForm = truth(
      rubric.free_form_criterion_comments
    );

    let body = paragraph(
      plain(rubric.title) || `Rubric ${rubric.id}`,
      { style: 'Title' }
    );

    if (
      !hidePoints &&
      !truth(rubric.hide_score_total) &&
      number(rubric.points_possible) != null
    ) {
      body += paragraph(
        `Total points: ${pointText(rubric.points_possible)}`,
        { keep: true }
      );
    }

    if (freeForm) {
      body += paragraph(
        'This rubric uses free-form comments.',
        { keep: true }
      );
    }

    if (!rubric.data.length) {
      body += paragraph(
        'No criteria have been added to this rubric.'
      );
    }

    rubric.data.forEach((criterion, index) => {
      body += paragraph(
        `${index + 1}. ${
          plain(criterion.description) || 'Criterion'
        }`,
        { style: 'Heading1' }
      );

      const detail = [];

      if (
        !hidePoints &&
        number(criterion.points) != null
      ) {
        detail.push(
          `Maximum: ${pointText(criterion.points)} points`
        );
      }

      if (truth(criterion.ignore_for_scoring)) {
        detail.push('Not counted towards the total score');
      }

      if (detail.length) {
        body += paragraph(detail.join('   |   '), {
          keep: true,
        });
      }

      const description = plain(
        criterion.long_description
      );

      if (description) {
        body += paragraph(description, { keep: true });
      }

      if (freeForm) {
        body += paragraph(
          'Comments: ____________________________________________________'
        );
      } else {
        const rows = ratingRows(
          criterion,
          rubric.rating_order,
          hidePoints
        );

        if (rows.length) {
          body += table(
            hidePoints
              ? ['Rating', 'Description']
              : ['Points', 'Rating', 'Description'],
            rows,
            hidePoints
              ? [2400, 6960]
              : [1360, 2200, 5800]
          );

          body += paragraph('', {
            style: 'TableGap',
          });
        } else {
          body += paragraph(
            'No ratings have been defined for this criterion.'
          );
        }
      }
    });

    const declaration =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

    const w =
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    const rel =
      'http://schemas.openxmlformats.org/package/2006/relationships';

    const officeRel =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    const documentXml =
      `${declaration}` +
      `<w:document xmlns:w="${w}"><w:body>${body}` +
      `<w:sectPr>` +
      `<w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1080" w:right="1440"` +
      ` w:bottom="1080" w:left="1440"` +
      ` w:header="360" w:footer="360" w:gutter="0"/>` +
      `</w:sectPr></w:body></w:document>`;

    const styles =
      `${declaration}<w:styles xmlns:w="${w}">` +

      `<w:docDefaults>` +
      `<w:rPrDefault><w:rPr>` +
      `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>` +
      `<w:color w:val="000000"/>` +
      `<w:sz w:val="22"/>` +
      `<w:lang w:val="en-AU"/>` +
      `</w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr>` +
      `<w:spacing w:after="120" w:line="264" w:lineRule="auto"/>` +
      `<w:widowControl/>` +
      `</w:pPr></w:pPrDefault>` +
      `</w:docDefaults>` +

      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">` +
      `<w:name w:val="Normal"/>` +
      `</w:style>` +

      `<w:style w:type="paragraph" w:styleId="Title">` +
      `<w:name w:val="Title"/>` +
      `<w:basedOn w:val="Normal"/>` +
      `<w:pPr><w:keepNext/>` +
      `<w:spacing w:after="200"/></w:pPr>` +
      `<w:rPr><w:b/><w:color w:val="000000"/>` +
      `<w:sz w:val="36"/></w:rPr>` +
      `</w:style>` +

      `<w:style w:type="paragraph" w:styleId="Heading1">` +
      `<w:name w:val="heading 1"/>` +
      `<w:basedOn w:val="Normal"/>` +
      `<w:pPr><w:keepNext/><w:keepLines/>` +
      `<w:spacing w:before="220" w:after="100"/>` +
      `<w:outlineLvl w:val="0"/></w:pPr>` +
      `<w:rPr><w:b/><w:color w:val="000000"/>` +
      `<w:sz w:val="26"/></w:rPr>` +
      `</w:style>` +

      `<w:style w:type="paragraph" w:styleId="TableText">` +
      `<w:name w:val="Table text"/>` +
      `<w:basedOn w:val="Normal"/>` +
      `<w:pPr>` +
      `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
      `</w:pPr></w:style>` +

      `<w:style w:type="paragraph" w:styleId="TableGap">` +
      `<w:name w:val="Table gap"/>` +
      `<w:basedOn w:val="Normal"/>` +
      `<w:pPr>` +
      `<w:spacing w:after="0" w:line="100" w:lineRule="exact"/>` +
      `</w:pPr><w:rPr><w:sz w:val="2"/></w:rPr>` +
      `</w:style></w:styles>`;

    return zip([
      [
        '[Content_Types].xml',
        `${declaration}` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels"` +
        ` ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml"` +
        ` ContentType="${MIME.replace(
          '.document',
          '.document.main+xml'
        )}"/>` +
        `<Override PartName="/word/styles.xml"` +
        ` ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
        `</Types>`,
      ],
      [
        '_rels/.rels',
        `${declaration}<Relationships xmlns="${rel}">` +
        `<Relationship Id="rId1"` +
        ` Type="${officeRel}/officeDocument"` +
        ` Target="word/document.xml"/>` +
        `</Relationships>`,
      ],
      [
        'word/document.xml',
        documentXml,
      ],
      [
        'word/styles.xml',
        styles,
      ],
      [
        'word/_rels/document.xml.rels',
        `${declaration}<Relationships xmlns="${rel}">` +
        `<Relationship Id="rId1"` +
        ` Type="${officeRel}/styles"` +
        ` Target="styles.xml"/>` +
        `</Relationships>`,
      ],
    ]);
  }

  // DOCX is a ZIP package of XML files.
  // This small ZIP writer avoids external dependencies.
  function zip(files) {
    const encoder = new TextEncoder();

    const crcTable = Array.from(
      { length: 256 },
      (_, n) => {
        let c = n;

        for (let k = 0; k < 8; k++) {
          c = c & 1
            ? 0xedb88320 ^ (c >>> 1)
            : c >>> 1;
        }

        return c >>> 0;
      }
    );

    function crc32(bytes) {
      let crc = 0xffffffff;

      for (const byte of bytes) {
        crc =
          crcTable[(crc ^ byte) & 255] ^
          (crc >>> 8);
      }

      return (crc ^ 0xffffffff) >>> 0;
    }

    function record(size) {
      const bytes = new Uint8Array(size);
      const view = new DataView(bytes.buffer);

      return {
        bytes,
        u16: (position, value) =>
          view.setUint16(position, value, true),
        u32: (position, value) =>
          view.setUint32(position, value, true),
      };
    }

    const localParts = [];
    const directory = [];
    let offset = 0;

    for (const [name, content] of files) {
      const nameBytes = encoder.encode(name);
      const data = encoder.encode(content);
      const crc = crc32(data);

      const local = record(30);

      local.u32(0, 0x04034b50);
      local.u16(4, 20);
      local.u16(6, 0x0800);
      local.u16(12, 0x0021);
      local.u32(14, crc);
      local.u32(18, data.length);
      local.u32(22, data.length);
      local.u16(26, nameBytes.length);

      localParts.push(
        local.bytes,
        nameBytes,
        data
      );

      const central = record(46);

      central.u32(0, 0x02014b50);
      central.u16(4, 20);
      central.u16(6, 20);
      central.u16(8, 0x0800);
      central.u16(14, 0x0021);
      central.u32(16, crc);
      central.u32(20, data.length);
      central.u32(24, data.length);
      central.u16(28, nameBytes.length);
      central.u32(42, offset);

      directory.push(
        central.bytes,
        nameBytes
      );

      offset +=
        30 + nameBytes.length + data.length;
    }

    const length = directory.reduce(
      (sum, bytes) => sum + bytes.length,
      0
    );

    const end = record(22);

    end.u32(0, 0x06054b50);
    end.u16(8, files.length);
    end.u16(10, files.length);
    end.u32(12, length);
    end.u32(16, offset);

    return new Blob(
      [...localParts, ...directory, end.bytes],
      { type: MIME }
    );
  }
})();
