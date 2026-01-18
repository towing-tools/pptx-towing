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
    }
    
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
    }
    .box.highlight { border-left-color: ${p.primary.secondary}; }
    .box.warning { border-left-color: ${p.functional.warning}; }
    .box.primary { background-color: ${p.primary.main}; color: ${p.functional.text_inverse}; border: none; }
    .box.primary h2 { color: ${p.primary.secondary}; }
    
    h2 {
        font-size: ${t.sizes.section_header};
        color: ${p.primary.secondary};
        margin-bottom: 4pt;
        font-family: ${t.font_family.heading};
        line-height: 1.2;
    }
    
    ul { padding-left: 1.2em; }
    li { margin-bottom: 3pt; }
    strong { color: ${p.functional.warning}; font-weight: bold; }
    blockquote { border-left: 3pt solid #ccc; padding-left: 10pt; font-style: italic; color: #666; }
    `;
}

function parseMarkdown(mdText) {
    const slides = mdText.split(/\n---\n/);

    return slides.map(slideMd => {
        let title = '';
        const lines = slideMd.split('\n');
        const htmlLines = [];
        const stack = []; // Track open containers: 'columns', 'column', 'box'
        let currentRatios = [1, 1];
        let columnIndex = 0; // Track which column we're in (0 = left, 1 = right)

        for (let i = 0; i < lines.length; i++) {
            let trimmed = lines[i].trim();
            if (!trimmed && !stack.includes('column')) continue;

            if (trimmed.startsWith('# ')) {
                title = trimmed.substring(2);
                continue;
            }

            // Start Columns container
            const colMatch = trimmed.match(/^::: columns *(\d+)\/(\d+)/);
            if (colMatch) {
                currentRatios = [parseInt(colMatch[1]), parseInt(colMatch[2])];
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
                // Start new column
                const ratio = columnIndex === 0 ? currentRatios[0] : currentRatios[1];
                htmlLines.push(`<div class="column" style="flex: ${ratio};">`);
                stack.push('column');
                continue;
            }

            // Start Box
            const boxMatch = trimmed.match(/^::: box *([a-z]*)/);
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
                let content = trimmed.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
                htmlLines.push(`<p>${content}</p>`);
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
    const slides = parseMarkdown(mdText);
    const css = generateCSS();
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    for (let i = 0; i < slides.length; i++) {
        const tempHtmlPath = path.join(path.dirname(outputPath), `slide_${i+1}.html`);
        fs.writeFileSync(tempHtmlPath, createSlideHtml(slides[i], css));
        await html2pptx(tempHtmlPath, pptx);
    }
    await pptx.writeFile({ fileName: outputPath });
}

run().catch(err => { console.error(err); process.exit(1); });