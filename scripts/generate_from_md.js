const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const html2pptx = require('./html2pptx');

// Load Configuration
const CONFIG_PATH = path.resolve(__dirname, '../assets/theme_config.json');
const CONFIG = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf-8'));

function generateCSS() {
    const p = CONFIG.palette;
    const t = CONFIG.typography;
    const l = CONFIG.layout;
    const c = CONFIG.components;

    return `
    body { 
        width: ${l.slide.width_pt}pt; 
        height: ${l.slide.height_pt}pt; 
        margin: 0; 
        padding: 0; 
        background-color: ${p.functional.background}; 
        font-family: ${t.font_family.base}; 
        overflow: hidden;
    }
    /* CSS Reset */
    h1, h2, p, ul, li { margin: 0; padding: 0; font-size: inherit; font-weight: inherit; }
    
    .slide-container { position: relative; width: 100%; height: 100%; }
    
    .slide-title {
        position: absolute;
        left: ${l.header.title.x_pt}pt;
        top: ${l.header.title.y_pt}pt;
        font-size: ${t.sizes.slide_title};
        font-weight: bold;
        color: ${p.primary.main};
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: calc(100% - ${l.header.title.x_pt * 2}pt);
    }
    /* 重要: タイトル色は #f4f0e8 固定。変更禁止 */
    
    .content {
        position: absolute;
        top: ${l.content.y_pt}pt;
        left: ${l.content.x_pt}pt;
        width: ${l.content.w_pt}pt;
        height: ${l.content.h_pt}pt;
        color: ${p.functional.text};
        font-size: ${t.sizes.body};
        line-height: ${t.line_height.base};
        display: flex;
        flex-direction: column;
        justify-content: center;
        gap: 12pt;
    }
    
    .columns { display: flex; gap: 20pt; width: 100%; align-items: flex-start; }
    
    .column { 
        display: flex; 
        flex-direction: column; 
        gap: 8pt; 
        min-width: 0; 
        flex-basis: 0; /* カラム幅を均等または比率通りに制限 */
        overflow-wrap: break-word;
    }
    
    .box {
        background-color: ${c.box.background_color};
        border-left: ${c.box.border_left_width_pt}pt solid ${p.functional.success};
        padding: ${c.box.padding_pt}pt;
        box-shadow: ${c.box.shadow};
        width: 100%;
        box-sizing: border-box;
        display: flex;
        flex-direction: column;
        gap: ${c.box.inner_gap_pt || 8}pt;
    }
    .box > * { margin: 0; }
    .box.highlight { border-left-color: ${p.primary.secondary}; }
    .box.warning { border-left-color: ${p.functional.warning}; }
    .box.primary { background-color: ${p.functional.success}; color: ${p.primary.main}; border: none; }
    .box.primary h2 { color: ${p.primary.main}; }
    .box.primary strong { color: ${p.primary.main}; text-decoration: underline; }
    /* 注意: .box.success と .box.green は使用禁止（見た目が悪いため削除済み） */

    /* KPI Dashboard styles */
    .box.kpi { text-align: center; padding: 16pt; }
    .box.kpi .number { font-size: 28pt; font-weight: bold; color: #4A4A3E; }
    .box.kpi .label { font-size: 11pt; color: #A3A099; margin-bottom: 4pt; }
    .box.kpi .badge { background-color: #1E8E3E; color: white; padding: 3pt 8pt; font-size: 10pt; display: inline-block; margin-top: 4pt; }

    /* Comparison table styles */
    .box.dark { background-color: #4A4A3E; color: ${p.functional.text_inverse}; border: none; }
    .box.gold { background-color: ${p.primary.secondary}; color: ${p.functional.text_inverse}; border: none; }
    .box.dark-gold { background-color: #766741; color: ${p.functional.text_inverse}; border: none; }
    .box.light-gray { background-color: #EFEEEB; color: ${p.functional.text}; border: none; }
    .box.gray { background-color: #A3A099; color: ${p.functional.text}; border: none; }

    /* Outline styles - DIVベースでPowerPoint形状に変換 */
    .outline-container {
        display: flex;
        flex-direction: column;
        gap: 12pt;
    }
    .outline-item {
        display: flex;
        align-items: flex-start;
        gap: 12pt;
    }
    .outline-number {
        width: 28pt;
        height: 28pt;
        background-color: ${p.primary.secondary};
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
    }
    .outline-number p {
        color: ${p.primary.main};
        font-weight: bold;
        font-size: 14pt;
        text-align: center;
        margin: 0;
    }
    .outline-text {
        flex: 1;
        padding-top: 4pt;
    }
    .outline-text p {
        color: ${p.functional.text};
        font-size: ${t.sizes.body};
        line-height: 1.4;
        margin: 0;
    }
    
    h2 {
        font-size: ${t.sizes.section_header};
        color: ${p.primary.secondary};
        margin-bottom: 4pt;
        font-family: ${t.font_family.heading};
        line-height: 1.2;
    }

    h3 {
        font-size: ${t.sizes.body};
        font-weight: bold;
        color: ${p.primary.secondary};
        margin-bottom: 4pt;
        font-family: ${t.font_family.heading};
        line-height: 1.2;
    }

    ul { padding-left: 1.2em; }
    li { margin-bottom: 3pt; }
    strong { color: ${p.functional.success}; font-weight: bold; }
    blockquote { border-left: 3pt solid #ccc; padding-left: 10pt; font-style: italic; color: #666; }

    /* Table styles - 表のスタイル */
    table {
        width: 100%;
        border-collapse: collapse;
        font-size: ${t.sizes.body};
        margin: 8pt 0;
    }
    table th {
        background-color: ${p.primary.secondary};
        color: ${p.primary.main};
        font-weight: bold;
        padding: 10pt 12pt;
        text-align: left;
        border: none;
    }
    table td {
        padding: 8pt 12pt;
        border-bottom: 1pt solid #E5E5E5;
        vertical-align: top;
    }
    table tr:nth-child(even) td {
        background-color: #F9F9F7;
    }
    table tr:hover td {
        background-color: #F4F0E8;
    }
    /* 表内の強調セル */
    table td.highlight {
        background-color: ${p.functional.success};
        color: ${p.primary.main};
        font-weight: bold;
    }
    table td.gold {
        background-color: ${p.primary.secondary};
        color: ${p.primary.main};
        font-weight: bold;
    }
    `;
}

function parseMarkdown(mdText) {
    // BOMを除去（存在する場合）
    if (mdText.charCodeAt(0) === 0xFEFF) {
        mdText = mdText.slice(1);
    }

    // 改行を統一（CRLFをLFに）
    mdText = mdText.replace(/\r\n/g, '\n');

    // ファイル先頭のYAMLフロントマター（***で囲まれた部分）を除去
    // 先頭の空白も考慮
    let cleanedMdText = mdText.replace(/^\s*\*\*\*[\s\S]*?\*\*\*\s*/, '');

    // ファイル先頭の---区切りも除去（YAML除去後に残っている場合）
    cleanedMdText = cleanedMdText.replace(/^\s*---\s*\n?/, '');

    // スライド分割: まず \n---\n で試みる
    let slides = cleanedMdText.split(/\n---\n/);

    // ---区切りがない場合（スライドが1つだけで、複数のYAMLブロックがある場合）
    // ***slide_number で始まるYAMLブロックをスライド区切りとして使用
    if (slides.length === 1 && cleanedMdText.includes('slide_number')) {
        // slide_number を含む *** ブロックの開始位置で分割
        const splitSlides = cleanedMdText.split(/\n(?=\*\*\*\s*\nslide_number)/);
        if (splitSlides.length > 1) {
            slides = splitSlides;
        }
    }

    return slides.map(slideMd => {
        // 先頭・末尾の空白を除去
        let cleanedSlideMd = slideMd.trim();
        // 各スライド先頭のYAMLメタデータ（***で囲まれた部分）を除去
        cleanedSlideMd = cleanedSlideMd.replace(/^\*\*\*[\s\S]*?\*\*\*\s*/, '');

        let title = '';
        const lines = cleanedSlideMd.split('\n');
        const htmlLines = [];
        const stack = []; // Track open containers: 'columns', 'column', 'box'
        let currentRatios = [1, 1];
        let columnIndex = 0; // Track which column we're in (0 = left, 1 = right)

        for (let i = 0; i < lines.length; i++) {
            let trimmed = lines[i].trim();
            if (!trimmed && !stack.includes('column')) continue;

            // h1: スライドタイトル
            if (trimmed.startsWith('# ')) {
                title = trimmed.substring(2);
                continue;
            }

            // h3: サブセクションヘッダー（ボックス内など）
            if (trimmed.startsWith('### ')) {
                let content = trimmed.substring(4).replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
                htmlLines.push(`<h3>${content}</h3>`);
                continue;
            }

            // Start Columns container (supports 2 or 3 columns)
            const colMatch = trimmed.match(/^::: columns *(\d+)\/(\d+)(?:\/(\d+))?/);
            if (colMatch) {
                if (colMatch[3]) {
                    // 3-column layout
                    currentRatios = [parseInt(colMatch[1]), parseInt(colMatch[2]), parseInt(colMatch[3])];
                } else {
                    // 2-column layout
                    currentRatios = [parseInt(colMatch[1]), parseInt(colMatch[2])];
                }
                columnIndex = 0;
                htmlLines.push('<div class="columns">');
                stack.push('columns');
                continue;
            }

            // Start individual Column
            if (trimmed === '::: column') {
                // Close previous column if exists
                if (stack[stack.length - 1] === 'column') {
                    htmlLines.push('</div>');
                    stack.pop();
                    columnIndex++;
                }
                // Start new column (supports 2 or 3 columns)
                const ratio = currentRatios[columnIndex] || currentRatios[currentRatios.length - 1];
                htmlLines.push(`<div class="column" style="flex: ${ratio};">`);
                stack.push('column');
                continue;
            }

            // Start Box (supports hyphenated class names like "dark-gold", "light-gray")
            const boxMatch = trimmed.match(/^::: box *([a-z-]*)/);
            if (boxMatch) {
                htmlLines.push(`<div class="box ${boxMatch[1]}">`);
                stack.push('box');
                continue;
            }

            // Close Tag :::
            if (trimmed === ':::') {
                const last = stack.pop();
                if (last === 'box') {
                    htmlLines.push('</div>');
                } else if (last === 'column') {
                    htmlLines.push('</div>');
                } else if (last === 'columns') {
                    htmlLines.push('</div>');
                    columnIndex = 0;
                }
                continue;
            }

            // Content Tags
            if (trimmed.startsWith('## ')) {
                htmlLines.push(`<h2>${trimmed.substring(3)}</h2>`);
            } else if (trimmed.startsWith('* ')) {
                // Simplified list handling
                const prevLine = htmlLines[htmlLines.length - 1];
                if (!prevLine?.includes('<li>') && !prevLine?.includes('<ul>')) {
                    htmlLines.push('<ul>');
                }
                let content = trimmed.substring(2).replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
                htmlLines.push(`<li>${content}</li>`);
                // Check if next line is not a list item, close <ul>
                const nextLine = lines[i + 1]?.trim();
                if (!nextLine?.startsWith('* ')) {
                    htmlLines.push('</ul>');
                }
            } else if (trimmed.startsWith('> ')) {
                htmlLines.push(`<blockquote>${trimmed.substring(2)}</blockquote>`);
            } else if (trimmed) {
                // HTMLタグで始まる行はそのまま出力（<p>で囲まない）
                if (trimmed.startsWith('<')) {
                    htmlLines.push(trimmed);
                } else {
                    let content = trimmed.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
                    htmlLines.push(`<p>${content}</p>`);
                }
            }
        }

        // Cleanup any unclosed tags
        while (stack.length > 0) {
            stack.pop();
            htmlLines.push('</div>');
        }

        return { title, html: htmlLines.join('\n') };
    });
}

function createSlideHtml(slideData, css) {
    return `<!DOCTYPE html><html><head><style>${css}</style></head>
    <body><div class="slide-container"><div class="slide-title"><h1>${slideData.title}</h1></div>
    <div class="content">${slideData.html}</div></div></body></html>`;
}

async function run() {
    const [inputPath, outputPath] = process.argv.slice(2);
    if (!inputPath || !outputPath) process.exit(1);

    const mdText = fs.readFileSync(inputPath, 'utf-8');
    const allSlides = parseMarkdown(mdText);

    // 空のスライド（タイトルも内容もないもの）をフィルタリング
    const slides = allSlides.filter(slide => slide.title || slide.html.trim());

    const css = generateCSS();
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    // 中間HTMLファイル用のslidesディレクトリを作成
    const outputDir = path.dirname(outputPath);
    const slidesDir = path.join(outputDir, 'slides');
    if (!fs.existsSync(slidesDir)) {
        fs.mkdirSync(slidesDir, { recursive: true });
    }

    // 古いHTMLファイルを削除
    const existingFiles = fs.readdirSync(slidesDir).filter(f => f.endsWith('.html'));
    existingFiles.forEach(f => fs.unlinkSync(path.join(slidesDir, f)));

    for (let i = 0; i < slides.length; i++) {
        // 中間HTMLファイルはslidesディレクトリに保存
        const tempHtmlPath = path.join(slidesDir, `slide_${i+1}.html`);
        fs.writeFileSync(tempHtmlPath, createSlideHtml(slides[i], css));
        await html2pptx(tempHtmlPath, pptx);
    }
    await pptx.writeFile({ fileName: outputPath });
    console.log(`✅ Generated: ${outputPath}`);
    console.log(`   Slides: ${slides.length}`);
    console.log(`   Intermediate HTML files saved in: ${slidesDir}`);
}

run().catch(err => { console.error(err); process.exit(1); });