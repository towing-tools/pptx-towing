const fs = require('fs');
const path = require('path');

// Parse XML and extract text content
function extractTextFromXml(xmlContent) {
    const texts = [];

    // Extract text from <a:t> tags (PowerPoint text elements)
    const textRegex = /<a:t>([^<]*)<\/a:t>/g;
    let match;
    while ((match = textRegex.exec(xmlContent)) !== null) {
        if (match[1].trim()) {
            texts.push(match[1].trim());
        }
    }

    return texts;
}

// Extract shape properties
function extractShapeInfo(xmlContent) {
    const shapes = [];

    // Find shape elements
    const shapeRegex = /<p:sp[\s\S]*?<\/p:sp>/g;
    let match;
    while ((match = shapeRegex.exec(xmlContent)) !== null) {
        const shape = match[0];
        const shapeInfo = {
            texts: extractTextFromXml(shape),
            hasTable: false
        };

        // Check for fill color
        const fillMatch = shape.match(/<a:srgbClr val="([A-Fa-f0-9]{6})"/);
        if (fillMatch) {
            shapeInfo.fillColor = '#' + fillMatch[1];
        }

        if (shapeInfo.texts.length > 0) {
            shapes.push(shapeInfo);
        }
    }

    // Find table elements
    const tableRegex = /<a:tbl[\s\S]*?<\/a:tbl>/g;
    while ((match = tableRegex.exec(xmlContent)) !== null) {
        const table = match[0];
        const rows = [];

        const rowRegex = /<a:tr[\s\S]*?<\/a:tr>/g;
        let rowMatch;
        while ((rowMatch = rowRegex.exec(table)) !== null) {
            const cells = extractTextFromXml(rowMatch[0]);
            if (cells.length > 0) {
                rows.push(cells);
            }
        }

        if (rows.length > 0) {
            shapes.push({
                isTable: true,
                rows: rows
            });
        }
    }

    return shapes;
}

function analyzeSlides(extractDir) {
    const slidesDir = path.join(extractDir, 'ppt', 'slides');
    const output = [];

    output.push('# sample_pptx.pptx デザイン分析');
    output.push('');
    output.push('## 概要');

    // Get all slide files
    const slideFiles = fs.readdirSync(slidesDir)
        .filter(f => f.match(/^slide\d+\.xml$/))
        .sort((a, b) => {
            const numA = parseInt(a.match(/\d+/)[0]);
            const numB = parseInt(b.match(/\d+/)[0]);
            return numA - numB;
        });

    output.push(`- スライド数: ${slideFiles.length}`);
    output.push('');

    for (const file of slideFiles) {
        const slideNum = parseInt(file.match(/\d+/)[0]);
        const xmlPath = path.join(slidesDir, file);
        const xmlContent = fs.readFileSync(xmlPath, 'utf-8');

        output.push('---');
        output.push('');
        output.push(`## スライド ${slideNum}`);
        output.push('');

        const shapes = extractShapeInfo(xmlContent);

        for (let i = 0; i < shapes.length; i++) {
            const shape = shapes[i];

            if (shape.isTable) {
                output.push('### 表');
                output.push('');
                for (const row of shape.rows) {
                    output.push('| ' + row.join(' | ') + ' |');
                }
                output.push('');
            } else if (shape.texts.length > 0) {
                output.push(`### テキスト要素 ${i + 1}`);
                if (shape.fillColor) {
                    output.push(`背景色: ${shape.fillColor}`);
                }
                output.push('');
                for (const text of shape.texts) {
                    output.push(`- ${text}`);
                }
                output.push('');
            }
        }
    }

    return output.join('\n');
}

// Main
const extractDir = process.argv[2] || './output/sample_analysis';
const outputFile = process.argv[3] || './output/sample_pptx_design.md';

try {
    const analysis = analyzeSlides(extractDir);
    fs.writeFileSync(outputFile, analysis, 'utf-8');
    console.log(`Analysis saved to: ${outputFile}`);
} catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
}
