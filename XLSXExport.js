const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs-extra');
const { toSystemPath } = require('../../../lib/core/path');

// ===================================================================
// COMPREHENSIVE XLSX EXPORT WITH ALL EXCELJS FEATURES
// Supports: Formatting, Fonts, Borders, Colors, Protection, 
// Validation, Conditional Formatting, Print Options, Freeze Panes
// ===================================================================

exports.xlsxexport = async function (options) {
    try {
        options = this.parse(options);
        let dirPath = this.parse(options.path);
        let data = options.data;
        let sheetname = (options.sheetname || 'Sheet1').substring(0, 31).replace(/[^a-zA-Z0-9 ]/g, '');
        let excelname = options.filename;
        let filetype = options.filetype || '.xlsx';

        // Validate inputs
        if (typeof dirPath !== 'string') {
            throw new Error('path: Path is required and must be a string.');
        }

        if (!data) {
            throw new Error('data: Data is required.');
        }

        if (!excelname) {
            throw new Error('filename: Filename is required.');
        }

        // Normalize file type
        if (!filetype.startsWith('.')) {
            filetype = '.' + filetype;
        }

        // Convert path to system path
        dirPath = toSystemPath(dirPath);
        console.log('System path:', dirPath);

        // Create full file path
        let fullPath = dirPath.endsWith(path.sep) ? dirPath : dirPath + path.sep;
        fullPath += excelname + filetype;
        console.log('Full path:', fullPath);

        // Ensure directory exists
        await fs.ensureDir(path.dirname(fullPath));


        // Create workbook and worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet(sheetname);

        // Convert single object to array
        const dataArray = Array.isArray(data) ? data : [data];

        if (dataArray.length === 0) {
            throw new Error('data: Data array cannot be empty.');
        }

        // Get headers from first object or use first array
        const isArrayOfObjects = typeof dataArray[0] === 'object' && !Array.isArray(dataArray[0]);
        const headers = isArrayOfObjects ? Object.keys(dataArray[0]) : dataArray[0];

        // ===== HEADER ROW =====
        const headerRow = worksheet.addRow(headers);
        headerRow.eachCell((cell) => {
            applyHeaderStyle(cell, options);
        });

        // ===== DATA ROWS =====
        dataArray.slice(1).forEach(item => {
            const rowData = isArrayOfObjects
                ? headers.map(header => item[header] || '')
                : item;
            const dataRow = worksheet.addRow(rowData);
            dataRow.eachCell((cell) => {
                applyDataCellStyle(cell, options);
            });
        });

        // ===== AUTO-FIT COLUMNS =====
        worksheet.columns.forEach((column, index) => {
            let maxLength = 0;
            column.eachCell?.({ includeEmpty: true }, (cell) => {
                const cellValue = cell.value ? String(cell.value) : '';
                maxLength = Math.max(maxLength, cellValue.length);
            });
            column.width = Math.min(Math.max(maxLength + 2, 10), 50);
        });

        // ===== SHEET PROTECTION =====
        if (options.protectSheet) {
            const protectOptions = {
                password: options.sheetPassword || '',
                sheet: true,
                content: true,
                objects: true,
                scenarios: true,
                formatRows: !options.lockCells,
                formatColumns: !options.lockCells,
                formatCells: false,
                insertRows: false,
                insertColumns: false,
                insertHyperlinks: false,
                deleteRows: false,
                deleteColumns: false,
                sort: false,
                autoFilter: false,
                pivotTables: false
            };
            worksheet.protect(options.sheetPassword || '', protectOptions);
        }

        // ===== FREEZE PANES =====
        if (options.freezePane && options.freezePane.trim()) {
            try {
                const [rows, cols] = options.freezePane.split(':').map(Number);
                if ((rows && rows > 0) || (cols && cols > 0)) {
                    worksheet.views = [{
                        state: 'frozen',
                        xSplit: cols || 0,
                        ySplit: rows || 0
                    }];
                }
            } catch (e) {
                console.warn('Freeze pane error:', e.message);
            }
        }

        // ===== AUTO FILTER =====
        if (options.autoFilter && options.autoFilter.trim()) {
            try {
                worksheet.autoFilter.from = options.autoFilter;
            } catch (e) {
                console.warn('Auto filter error:', e.message);
            }
        }

        // ===== CONDITIONAL FORMATTING =====
        if (options.conditionalFormatting && options.conditionalFormatting.trim()) {
            try {
                const cfRules = typeof options.conditionalFormatting === 'string'
                    ? JSON.parse(options.conditionalFormatting)
                    : options.conditionalFormatting;

                if (Array.isArray(cfRules)) {
                    cfRules.forEach(rule => {
                        if (rule.range && rule.type && rule.formula) {
                            worksheet.addConditionalFormatting({
                                ref: rule.range,
                                rules: [{
                                    type: rule.type,
                                    formulae: [rule.formula],
                                    priority: rule.priority || 1,
                                    stopIfTrue: rule.stopIfTrue || false,
                                    style: {
                                        font: rule.font || {},
                                        fill: rule.fill || {},
                                        border: rule.border || {}
                                    }
                                }]
                            });
                        }
                    });
                }
            } catch (e) {
                console.warn('Conditional formatting error:', e.message);
            }
        }

        // ===== DATA VALIDATION =====
        if (options.dataValidation && options.dataValidation.trim()) {
            try {
                const validationRules = typeof options.dataValidation === 'string'
                    ? JSON.parse(options.dataValidation)
                    : options.dataValidation;

                if (Array.isArray(validationRules)) {
                    validationRules.forEach(rule => {
                        if (rule.range && rule.type) {
                            const validation = {
                                type: rule.type, // 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'
                                operator: rule.operator || 'equal',
                                formula1: rule.formula1,
                                formula2: rule.formula2,
                                showInputMessage: rule.showInputMessage !== false,
                                showErrorMessage: rule.showErrorMessage !== false,
                                errorTitle: rule.errorTitle || 'Invalid Entry',
                                error: rule.error || 'Invalid value',
                                promptTitle: rule.promptTitle,
                                prompt: rule.prompt,
                                allow: rule.allow // 'list', 'whole', 'decimal', etc.
                            };
                            worksheet.dataValidations.add(rule.range, validation);
                        }
                    });
                }
            } catch (e) {
                console.warn('Data validation error:', e.message);
            }
        }

        // ===== PRINT OPTIONS =====
        if (options.printOptions && options.printOptions.trim()) {
            try {
                const printOpts = typeof options.printOptions === 'string'
                    ? JSON.parse(options.printOptions)
                    : options.printOptions;

                if (printOpts.orientation) worksheet.pageSetup.orientation = printOpts.orientation;
                if (printOpts.paperSize) worksheet.pageSetup.paperSize = printOpts.paperSize;
                if (printOpts.margins) worksheet.pageMargins = printOpts.margins;
                if (printOpts.fitToPage) {
                    worksheet.pageSetup.fitToPage = true;
                    worksheet.pageSetup.fitToHeight = printOpts.fitToHeight || 1;
                    worksheet.pageSetup.fitToWidth = printOpts.fitToWidth || 1;
                }
                if (printOpts.printGridLines) worksheet.pageSetup.printGridLines = true;
                if (printOpts.showRowColHeaders) worksheet.pageSetup.showRowColHeaders = true;
            } catch (e) {
                console.warn('Print options error:', e.message);
            }
        }

        // ===== WRITE FILE =====
        try {
            console.log('Attempting to write file to:', fullPath);
            await workbook.xlsx.writeFile(fullPath);

            // Verify file was created
            const fileExists = await fs.pathExists(fullPath);
            if (!fileExists) {
                throw new Error(`File write completed but file not found at ${fullPath}`);
            }

            const stats = await fs.stat(fullPath);
            console.log(`✓ File created successfully: ${fullPath} (${stats.size} bytes)`);
        } catch (writeError) {
            throw new Error(`Failed to write Excel file: ${writeError.message}`);
        }

        // Return relative path for client access
        const FilePath = '/' + path.join(path.basename(dirPath), excelname + filetype).replace(/\\/g, '/');
        return { 
            FilePath: FilePath,
            FileName: excelname + filetype,
            Success: true
        };

    } catch (error) {
        console.error('XLSX Export Error:', error.message);
        console.error('Stack:', error.stack);
        throw new Error(`XLSX Export failed: ${error.message}`);
    }
};

/**
 * Apply comprehensive header cell styling
 */
function applyHeaderStyle(cell, options) {
    // Alignment
    cell.alignment = {
        horizontal: options.alignment || 'center',
        vertical: options.verticalAlignment || 'middle',
        wrapText: options.wrapText === true,
        indent: 0,
        rtl: false
    };

    // Font styling
    cell.font = {
        name: options.fontName || 'Arial',
        size: Math.max((options.fontSize || 11) + 1, 11),
        bold: true,
        italic: options.fontItalic === true,
        color: { argb: options.fontColor || 'FF000000' },
        underline: options.fontUnderline && options.fontUnderline !== 'none' ? options.fontUnderline : undefined,
        strike: false,
        outline: false,
        shadow: false,
        vertAlign: 'superscript'
    };

    // Fill/Background
    const fillColor = options.backgroundColor || 'FFD3D3D3';
    if (options.fillType === 'gradient') {
        // Gradient fill requires stops array
        cell.fill = {
            type: 'gradient',
            gradient: 'angle',
            angle: options.gradientAngle || 0,
            stops: [
                { position: 0, color: { argb: options.gradientStartColor || 'FFFFFFFF' } },
                { position: 100, color: { argb: fillColor } }
            ]
        };
    } else {
        // Pattern fill
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: fillColor },
            bgColor: { argb: 'FFFFFFFF' }
        };
    }

    // Border
    const borderStyle = options.borderStyle || 'thin';
    const borderColor = options.borderColor || 'FF000000';
    cell.border = {
        left: { style: borderStyle, color: { argb: borderColor } },
        right: { style: borderStyle, color: { argb: borderColor } },
        top: { style: borderStyle, color: { argb: borderColor } },
        bottom: { style: borderStyle, color: { argb: borderColor } },
        diagonal: { up: false, down: false }
    };

    // Protection
    cell.protection = {
        locked: options.lockCells === true,
        hidden: options.lockCells === true
    };

    // Number format for headers
    cell.numFmt = '@'; // Text format for headers
}

/**
 * Apply comprehensive data cell styling
 */
function applyDataCellStyle(cell, options) {
    // Number formatting
    if (options.numberFormat && options.numberFormat.trim()) {
        cell.numFmt = options.numberFormat;
    }

    // Alignment
    cell.alignment = {
        horizontal: options.alignment || 'left',
        vertical: options.verticalAlignment || 'top',
        wrapText: options.wrapText === true,
        indent: 0,
        rtl: false
    };

    // Font styling
    cell.font = {
        name: options.fontName || 'Arial',
        size: options.fontSize || 11,
        bold: options.fontBold === true,
        italic: options.fontItalic === true,
        color: { argb: options.fontColor || 'FF000000' },
        underline: options.fontUnderline && options.fontUnderline !== 'none' ? options.fontUnderline : undefined,
        strike: false,
        outline: false,
        shadow: false
    };

    // Fill/Background
    if (options.fillType !== 'none') {
        const fillColor = options.backgroundColor || 'FFFFFFFF';
        if (options.fillType === 'gradient') {
            // Gradient fill requires stops array
            cell.fill = {
                type: 'gradient',
                gradient: 'angle',
                angle: options.gradientAngle || 0,
                stops: [
                    { position: 0, color: { argb: options.gradientStartColor || 'FFFFFFFF' } },
                    { position: 100, color: { argb: fillColor } }
                ]
            };
        } else {
            // Pattern fill
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: fillColor },
                bgColor: { argb: 'FFFFFFFF' }
            };
        }
    }

    // Border
    if (options.borderStyle !== 'none') {
        const borderStyle = options.borderStyle;
        const borderColor = options.borderColor || 'FF000000';
        cell.border = {
            left: { style: borderStyle, color: { argb: borderColor } },
            right: { style: borderStyle, color: { argb: borderColor } },
            top: { style: borderStyle, color: { argb: borderColor } },
            bottom: { style: borderStyle, color: { argb: borderColor } },
            diagonal: { up: false, down: false }
        };
    }

    // Protection
    cell.protection = {
        locked: options.lockCells === true,
        hidden: options.lockCells === true
    };
}
