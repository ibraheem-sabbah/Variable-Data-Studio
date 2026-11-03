#target illustrator
/*
  Variable Data Studio for Adobe Illustrator
  Version: 2.0.7

  Generation hardening: Separate Artboards now reports the record and stage that fails. Text
  replacement treats Illustrator character/paragraph collections defensively and keeps the
  duplicated frame's native style if an individual collection entry becomes stale.

  UI rebuild: the mapping list uses exactly five permanent ScriptUI row slots bound to a separate
  variable data model. Scrolling rebinds these five controls instead of hiding/removing dozens of
  controls, preventing off-screen controls from affecting layout and keeping the viewport stable.

  Workflow:
  1) Select a source artwork Group, or restore it from a saved Variable Data Studio preset when no Group is selected.
  2) Assign custom-named variables to live text frames, or load a saved preset. The mapping list shows five rows and scrolls vertically.
  3) Enter/edit records in the paged grid, paste Excel data, or import a selected workbook worksheet.
  4) Save presets with settings/mappings even when the record table is empty; presets can be exported and reloaded.
  5) When saving a preset, the script also attempts to save a standalone Illustrator copy of the source output.

  Notes:
  - This script uses Illustrator ExtendScript (ES3-compatible syntax).
  - Distances are entered in centimeters and converted to points internally.
  - Separate Artboards reuse the dimensions of the artboard containing the source output.
  - Static artwork can optionally be rasterized while mapped text stays editable.
  - Excel workbooks are imported through Microsoft Excel COM on Windows; CSV/TSV/TXT can be read directly.
*/

(function () {
    var SCRIPT_NAME = "Variable Data Studio";
    var SCRIPT_VERSION = "2.0.7";
    var PT_PER_CM = 72 / 2.54;
    var MAX_ARTBOARDS = 100;

    var CHARACTER_ATTRIBUTE_NAMES = [
        "akiLeft", "akiRight", "alignment", "alternateGlyphs", "autoLeading", "baselineDirection",
        "baselinePosition", "baselineShift", "capitalization", "connectionForms", "contextualLigature",
        "discretionaryLigature", "figureStyle", "fillColor", "fractions", "horizontalScale", "italics",
        "kerningMethod", "language", "leading", "ligature", "noBreak", "openTypePosition", "ordinals",
        "ornaments", "overprintFill", "overprintStroke", "proportionalMetrics", "rotation", "size",
        "strikeThrough", "stylisticAlternates", "swash",
        "tateChuYokoHorizontal", "tateChuYokoVertical", "textFont", "titling", "tracking", "Tsume",
        "underline", "verticalScale", "wariChuCharactersAfterBreak", "wariChuCharactersBeforeBreak",
        "waiChuEnabled", "wariChuJustification", "wariChuLineGap", "wariChuLines", "wariChuScale"
    ];
    var PARAGRAPH_ATTRIBUTE_NAMES = [
        "autoHyphenate", "burasagari", "desiredGlyphScaling", "desiredLetterSpacing",
        "desiredWordSpacing", "everyLineComposer", "firstLineIndent", "hyphenateCapitalizedWords",
        "hyphenation", "hyphenationZone", "hyphenLimit", "justification", "keepAllLinesTogether",
        "keepFirstLines", "keepLastLines", "kinsoku", "kurikaeshiMojiShori", "leadingType", "leftIndent",
        "maximumGlyphScaling", "maximumLetterSpacing", "maximumWordSpacing", "minimumGlyphScaling",
        "minimumLetterSpacing", "minimumWordSpacing", "paragraphDirection", "rightIndent", "romanHanging",
        "singleWordJustification", "spaceAfter", "spaceBefore", "tabStops", "wordSpacing"
    ];

    var doc = app.documents.length > 0 ? app.activeDocument : null;
    var template = doc ? getSelectedTemplate(doc) : null;
    var startupPreset = null;

    // If no Group is selected, allow the user to restore the saved template from a Variable Data Studio preset.
    if (!template) {
        var tryPresetRestore = confirm(
            "No source artwork Group is selected.\n\nLoad a saved Variable Data Studio preset and restore its saved template?\nChoose No to cancel and select a Group manually.",
            true,
            SCRIPT_NAME
        );
        if (!tryPresetRestore) {
            alert("Open a document, group the complete artwork design, select that Group, and run the script again.", SCRIPT_NAME);
            return;
        }

        var startupPresetFile = choosePresetFile("Select a Variable Data Studio preset (.vds or legacy .cbg)");
        if (!startupPresetFile) { return; }
        try {
            startupPreset = readPresetFile(startupPresetFile);
            rememberLastPresetPath(startupPresetFile);
        } catch (startupPresetError) {
            alert("The preset could not be loaded.\n\n" + errorMessage(startupPresetError), SCRIPT_NAME + " - Load Preset");
            return;
        }

        // Prefer the exact template backup saved with this preset. This avoids accidentally
        // using an unrelated Group in the currently active document with a similar name.
        var backupFile = resolvePresetTemplateFile(startupPresetFile, startupPreset, false);
        if (backupFile) {
            try {
                doc = app.open(backupFile);
                template = findPresetTemplateGroup(doc, startupPreset);
            } catch (openTemplateError) {
                alert("Could not open the saved template file.\n\n" + errorMessage(openTemplateError), SCRIPT_NAME + " - Restore Template");
                return;
            }
        }
        // If there is no stored backup, a matching Group already open may still be used.
        if (!template && doc) { template = findPresetTemplateGroup(doc, startupPreset); }
        // Last resort: let the user locate the matching Illustrator template manually.
        if (!template) {
            backupFile = resolvePresetTemplateFile(startupPresetFile, startupPreset, true);
            if (!backupFile) {
                alert(
                    "Preset settings were read, but no template Group could be opened.\n\n" +
                    "Make sure the preset's _Template.ai file is available, or choose the matching Illustrator template file.",
                    SCRIPT_NAME + " - Restore Template"
                );
                return;
            }
            try {
                doc = app.open(backupFile);
                template = findPresetTemplateGroup(doc, startupPreset);
            } catch (openManualTemplateError) {
                alert("Could not open the selected template file.\n\n" + errorMessage(openManualTemplateError), SCRIPT_NAME + " - Restore Template");
                return;
            }
        }
        if (!doc || !template) {
            alert(
                "The selected preset's template Group could not be found.\n\n" +
                "Open the saved _Template.ai file, select the full source artwork Group, and run the script again.",
                SCRIPT_NAME + " - Restore Template"
            );
            return;
        }
        try { app.activeDocument = doc; doc.selection = []; template.selected = true; } catch (selectRestoredTemplateError) {}
    }

    var templateFrames = collectTextFrames(template);
    if (templateFrames.length === 0) {
        alert(
            "The selected Group contains no editable text frames.\n\n" +
            "At least one live-text object is required. Only text frames can be assigned to variables.",
            SCRIPT_NAME
        );
        return;
    }

    var originalBounds;
    try {
        originalBounds = getArtworkBounds(template);
    } catch (boundsError) {
        alert("Could not read the selected Group's bounds.\n\n" + errorMessage(boundsError), SCRIPT_NAME);
        return;
    }

    var templateTextSnapshot = snapshotTextFrames(templateFrames);
    var templateOriginalName = "";
    try { templateOriginalName = String(template.name); } catch (ignoreOriginalName) { templateOriginalName = ""; }
    var originalActiveArtboardIndex = doc.artboards.getActiveArtboardIndex();
    var originalArtboardCount = doc.artboards.length;
    var originalActiveArtboardRect = copyRect(doc.artboards[originalActiveArtboardIndex].artboardRect);
    var sourceArtboardIndex = findContainingArtboardIndex(doc, originalBounds, originalActiveArtboardIndex);
    var sourceArtboardRect = copyRect(doc.artboards[sourceArtboardIndex].artboardRect);

    var wizardResult = runTwoStageWizard(templateFrames, template.name, doc, template, sourceArtboardRect, originalBounds, startupPreset);
    if (!wizardResult || wizardResult.saveOnly) {
        return;
    }

    var config = wizardResult.config;
    config.blankBehavior = wizardResult.blankBehavior || "Clear mapped text";
    var parsedData;
    try {
        parsedData = makeDataRowsFromGrid(wizardResult.records, config.variables);
        validateConfiguration(config, parsedData.rows.length);
    } catch (validationError) {
        alert(errorMessage(validationError), SCRIPT_NAME + " - Validation");
        return;
    }

    if (config.layout === "Same Position") {
        if (!confirm(
            "Same Position will stack all outputs exactly on top of one another.\n\n" +
            "Only the topmost output will be fully visible in the artwork.\n\n" +
            "Continue?",
            true,
            SCRIPT_NAME
        )) {
            return;
        }
    }

    if (config.layout === "Separate Artboards" && originalArtboardCount + parsedData.rows.length > MAX_ARTBOARDS) {
        alert(
            "This document would exceed Illustrator's limit of " + MAX_ARTBOARDS + " artboards.\n\n" +
            "Current artboards: " + originalArtboardCount + "\n" +
            "Requested output artboards: " + parsedData.rows.length + "\n\n" +
            "Reduce the number of outputs or use another layout.",
            SCRIPT_NAME
        );
        return;
    }

    var createdItems = [];
    var createdArtboardIndices = [];
    var outputItems = [];
    var artboardWasResized = false;
    var sourceTextWasChanged = false;
    var sourceWasMoved = false;
    var sourceMoveDelta = [0, 0];
    var sourceWasRemoved = false;
    var operationCommitted = false;
    var postSuccessWarning = "";

    try {
        if (config.layout === "Separate Artboards") {
            outputItems = generateOnSeparateArtboards(
                doc,
                template,
                config.variables,
                config.frameIndices,
                parsedData.rows,
                originalBounds,
                originalArtboardCount,
                sourceArtboardRect,
                config,
                createdItems,
                createdArtboardIndices
            );

            if (config.action === "Replace Original" && !config.flattenStaticArtwork) {
                // Preserve the source Group as record #1 and move it by the same offset as its artboard.
                sourceTextWasChanged = true;
                replacePlaceholders(template, config.variables, config.frameIndices, parsedData.rows[0], config.blankBehavior, 1);
                var firstBoardRect = doc.artboards[originalArtboardCount].artboardRect;
                sourceMoveDelta = [firstBoardRect[0] - sourceArtboardRect[0], firstBoardRect[1] - sourceArtboardRect[1]];
                if (sourceMoveDelta[0] !== 0 || sourceMoveDelta[1] !== 0) {
                    sourceWasMoved = true;
                    template.translate(sourceMoveDelta[0], sourceMoveDelta[1]);
                }
                try { template.name = makeOutputName(1); } catch (ignoreName0) {}
                try { createdItems[0].remove(); } catch (removeStagedFirstError) {
                    throw new Error("Could not finalize the first output on its dedicated artboard: " + errorMessage(removeStagedFirstError));
                }
                outputItems[0] = template;
            } else if (config.action === "Delete Original" || (config.action === "Replace Original" && config.flattenStaticArtwork)) {
                template.remove();
                sourceWasRemoved = true;
            }
        } else {
            outputItems = generateOnCanvas(
                doc,
                template,
                config.variables,
                config.frameIndices,
                parsedData.rows,
                originalBounds,
                config,
                createdItems
            );

            if (config.action === "Replace Original" && !config.flattenStaticArtwork) {
                sourceTextWasChanged = true;
                replacePlaceholders(template, config.variables, config.frameIndices, parsedData.rows[0], config.blankBehavior, 1);
                try { template.name = makeOutputName(1); } catch (ignoreName1) {}
                outputItems[0] = template;
            }

            // Grid/Horizontal/Vertical/Same Position keep the original artboard dimensions unchanged.
            // Generated items may extend beyond the artboard; the user's artboard is never resized here.

            if (config.action === "Delete Original" || (config.action === "Replace Original" && config.flattenStaticArtwork)) {
                template.remove();
                sourceWasRemoved = true;
            }
        }

        operationCommitted = true;
    } catch (generationError) {
        // Roll back any source changes, created artwork, artboards, and artboard-size changes.
        if (!sourceWasRemoved) {
            if (sourceWasMoved) {
                try { template.translate(-sourceMoveDelta[0], -sourceMoveDelta[1]); } catch (restoreMoveError) {}
            }
            if (sourceTextWasChanged) {
                restoreTextSnapshot(templateTextSnapshot);
            }
            try { template.name = templateOriginalName; } catch (restoreNameError) {}
        }

        if (artboardWasResized) {
            try { doc.artboards[originalActiveArtboardIndex].artboardRect = originalActiveArtboardRect; } catch (restoreBoardError) {}
        }

        removeCreatedItems(createdItems);
        removeCreatedArtboards(doc, createdArtboardIndices);

        alert(
            "The outputs could not be generated. The script attempted to roll back its changes.\n\n" +
            errorMessage(generationError),
            SCRIPT_NAME + " - Error"
        );
        return;
    }

    if (operationCommitted && outputItems.length > 0) {
        var requestedPdfMode = wizardResult.pdfMode || "prompt";
        var shouldExportSeparate = requestedPdfMode !== "single" && confirm(
            "Export each generated output as a separate, numbered PDF for printing?\n\nChoose No to skip PDF export.",
            true,
            SCRIPT_NAME + " - PDF Export"
        );
        if (requestedPdfMode === "single") {
            try {
                var singlePdfResult = exportOutputsAsSinglePDF(doc, outputItems, sourceArtboardRect, originalBounds);
                if (singlePdfResult && singlePdfResult.file) {
                    postSuccessWarning += "\n\nCombined PDF export: " + singlePdfResult.count + " page(s) saved to:\n" + singlePdfResult.file;
                }
            } catch (singlePdfError) {
                postSuccessWarning += "\n\nCombined PDF export did not complete: " + errorMessage(singlePdfError);
            }
        } else if (shouldExportSeparate) {
            try {
                var pdfExportResult = exportOutputsAsNumberedPDFs(doc, outputItems, sourceArtboardRect, originalBounds);
                if (pdfExportResult && pdfExportResult.count > 0) {
                    postSuccessWarning += "\n\nPDF export: " + pdfExportResult.count + " numbered file(s) saved to:\n" + pdfExportResult.folder + "\nFilename pattern: " + pdfExportResult.prefix + "_001.pdf";
                }
            } catch (pdfExportError) {
                postSuccessWarning += "\n\nPDF export did not complete: " + errorMessage(pdfExportError);
            }
        }
    }

    if (operationCommitted) {
        selectOutputItems(doc, outputItems);
        if (config.layout === "Separate Artboards") {
            try { doc.artboards.setActiveArtboardIndex(originalArtboardCount); } catch (setActiveError) {}
        }

        var summary = "Successfully generated " + parsedData.rows.length + " output item(s).\n\n" +
            "Layout: " + config.layout + "\n" +
            "Template: " + config.action + "\n" +
            "Variables: " + config.variables.join(", ");

        if (config.layout === "Separate Artboards") {
            summary += "\n\nEach output has its own artboard, using the same width and height as the source output artboard.";
            summary += "\nArtboard arrangement: " + (config.artboardArrangement || "Grid");
            if (config.action === "Replace Original" && !config.flattenStaticArtwork) {
                summary += "\nThe original Group was used for record #1 and moved onto the first new output artboard.";
            } else if (config.action === "Replace Original" && config.flattenStaticArtwork) {
                summary += "\nThe original Group was replaced by the flattened first output.";
            }
        } else {
            summary += "\n\nThe existing artboard dimensions were kept unchanged.";
        }
        if (config.flattenStaticArtwork) {
            summary += "\n\nStatic artwork was rasterized at " + config.rasterResolution + " ppi. Mapped variable text remains editable.";
        }
        summary += "\n\nThis Script Made By Ibraheem Sabbah - www.linkedin.com/in/ibraheem-sabbah";
        alert(summary, SCRIPT_NAME + " " + SCRIPT_VERSION);
    }

    function duplicatePageItemIntoDocument(item, sourceDocument, targetDocument) {
        var duplicate = null;
        try {
            duplicate = item.duplicate(targetDocument.layers[0], ElementPlacement.PLACEATEND);
            return duplicate;
        } catch (directDuplicateError) {
            // Clipboard fallback for Illustrator versions that reject cross-document duplicate().
            try {
                app.activeDocument = sourceDocument;
                sourceDocument.selection = [];
                item.selected = true;
                app.executeMenuCommand("copy");
                app.activeDocument = targetDocument;
                app.executeMenuCommand("pasteInPlace");
                if (targetDocument.selection && targetDocument.selection.length > 0) {
                    duplicate = targetDocument.selection[0];
                    targetDocument.selection = [];
                    return duplicate;
                }
            } catch (clipboardDuplicateError) {}
            throw new Error("Illustrator could not copy the output into a temporary export document. " + errorMessage(directDuplicateError));
        }
    }

    function makeTemporaryOutputDocument(sourceDocument, sourceItem, sourceArtboardRect, sourceBounds, variableNames, frameIndices, record, blankBehavior) {
        var width = Number(sourceArtboardRect[2]) - Number(sourceArtboardRect[0]);
        var height = Number(sourceArtboardRect[1]) - Number(sourceArtboardRect[3]);
        if (!(width > 0) || !(height > 0)) { throw new Error("The source artboard has invalid dimensions for preview/PDF output."); }
        var temporaryDocument = app.documents.add(sourceDocument.documentColorSpace, width, height);
        try {
            temporaryDocument.artboards[0].artboardRect = [0, height, width, 0];
            var duplicate = duplicatePageItemIntoDocument(sourceItem, sourceDocument, temporaryDocument);
            if (record && variableNames && frameIndices) {
                replacePlaceholders(duplicate, variableNames, frameIndices, record, blankBehavior || "Clear mapped text", 1);
            }
            var duplicateBounds = getArtworkBounds(duplicate);
            // Preserve the output's position relative to the source artboard, regardless of its layout position.
            var targetLeft = sourceBounds[0] - sourceArtboardRect[0];
            var targetTop = height - (sourceArtboardRect[1] - sourceBounds[1]);
            duplicate.translate(targetLeft - duplicateBounds[0], targetTop - duplicateBounds[1]);
            // Fit only if the artwork exceeds the source page. Do not add an artificial inset margin.
            fitPageItemWithinArtboard(duplicate, temporaryDocument.artboards[0].artboardRect, 0.0);
            app.activeDocument = temporaryDocument;
            return temporaryDocument;
        } catch (temporaryDocumentError) {
            try { temporaryDocument.close(SaveOptions.DONOTSAVECHANGES); } catch (ignoreTempDocClose) {}
            try { app.activeDocument = sourceDocument; } catch (ignoreRestoreSourceDoc) {}
            throw temporaryDocumentError;
        }
    }

    function fitPageItemWithinArtboard(item, artboardRect, marginRatio) {
        if (!item || !artboardRect || artboardRect.length < 4) {
            throw new Error("Could not fit the artwork because the page or artwork reference is missing.");
        }
        var pageLeft = Number(artboardRect[0]);
        var pageTop = Number(artboardRect[1]);
        var pageRight = Number(artboardRect[2]);
        var pageBottom = Number(artboardRect[3]);
        var pageWidth = pageRight - pageLeft;
        var pageHeight = pageTop - pageBottom;
        if (!(pageWidth > 0) || !(pageHeight > 0)) {
            throw new Error("The target page has invalid dimensions.");
        }
        var margin = Math.max(0, Math.min(0.08, Number(marginRatio) || 0));
        var bounds = getArtworkBounds(item);
        var itemWidth = Number(bounds[2]) - Number(bounds[0]);
        var itemHeight = Number(bounds[1]) - Number(bounds[3]);
        if (!(itemWidth > 0) || !(itemHeight > 0)) {
            throw new Error("The artwork has invalid visible bounds and cannot be fitted to the page.");
        }
        var availableWidth = pageWidth * (1 - margin * 2);
        var availableHeight = pageHeight * (1 - margin * 2);
        var scale = Math.min(availableWidth / itemWidth, availableHeight / itemHeight, 1);
        if (scale < 0.999999) {
            try {
                item.resize(scale * 100, scale * 100, true, true, true, true, true, Transformation.CENTER);
            } catch (resizeWithCenterError) {
                item.resize(scale * 100, scale * 100, true, true, true, true, true);
            }
            bounds = getArtworkBounds(item);
        }
        var centerX = (Number(bounds[0]) + Number(bounds[2])) / 2;
        var centerY = (Number(bounds[1]) + Number(bounds[3])) / 2;
        var pageCenterX = (pageLeft + pageRight) / 2;
        var pageCenterY = (pageTop + pageBottom) / 2;
        item.translate(pageCenterX - centerX, pageCenterY - centerY);
    }

    function showFirstRecordPreview(context, variableNames, firstRecord, blankBehavior) {
        if (!context || !context.documentRef || !context.templateGroup) {
            throw new Error("The original output template is not available for preview.");
        }
        var sourceDocument = context.documentRef;
        var sourceItem = context.templateGroup;
        var sourceArtboardRect = context.sourceArtboardRect;
        var sourceBounds = context.sourceBounds || getArtworkBounds(sourceItem);
        var temporaryDocument = null;
        var previewFile = new File(Folder.temp.fsName + "/VDS_preview_" + new Date().getTime() + ".png");
        var currentDocument = app.activeDocument;
        try {
            temporaryDocument = makeTemporaryOutputDocument(
                sourceDocument, sourceItem, sourceArtboardRect, sourceBounds,
                variableNames, context.config.frameIndices, firstRecord, blankBehavior
            );
            var pngOptions = new ExportOptionsPNG24();
            pngOptions.antiAliasing = true;
            pngOptions.transparency = true;
            pngOptions.artBoardClipping = true;
            try {
                var previewRect = temporaryDocument.artboards[0].artboardRect;
                var previewWidth = Math.abs(previewRect[2] - previewRect[0]);
                var previewHeight = Math.abs(previewRect[1] - previewRect[3]);
                var previewScale = Math.min(100, (680 * 100 / Math.max(1, previewWidth)), (420 * 100 / Math.max(1, previewHeight)));
                pngOptions.horizontalScale = previewScale;
                pngOptions.verticalScale = previewScale;
            } catch (ignorePngScale) {}
            temporaryDocument.exportFile(previewFile, ExportType.PNG24, pngOptions);
            temporaryDocument.close(SaveOptions.DONOTSAVECHANGES);
            temporaryDocument = null;
            app.activeDocument = sourceDocument;
            if (!previewFile.exists) { throw new Error("Illustrator did not create the temporary preview image."); }

            var previewWindow = new Window("dialog", SCRIPT_NAME + " - First Record Preview");
            previewWindow.orientation = "column";
            previewWindow.alignChildren = ["fill", "top"];
            previewWindow.spacing = 6;
            previewWindow.margins = 10;
            previewWindow.preferredSize = [720, 540];
            previewWindow.add("statictext", undefined, "Preview of record #1: mapped text is populated using the first grid row.");
            var previewImage = previewWindow.add("image", undefined, ScriptUI.newImage(previewFile));
            previewImage.alignment = ["center", "center"];
            previewImage.preferredSize = [690, 460];
            previewImage.maximumSize = [690, 460];
            var closeRow = previewWindow.add("group");
            closeRow.alignment = ["right", "top"];
            closeRow.add("button", undefined, "Close", { name: "ok" });
            previewWindow.show();
        } finally {
            try { if (temporaryDocument) { temporaryDocument.close(SaveOptions.DONOTSAVECHANGES); } } catch (ignorePreviewDocClose) {}
            try { app.activeDocument = currentDocument; } catch (ignorePreviewRestore) {}
            try { if (previewFile.exists) { previewFile.remove(); } } catch (ignorePreviewFileRemove) {}
        }
    }

    // PDFSaveOptions enumeration names vary in casing across Illustrator scripting references
    // and releases. Resolve enum members defensively so an unsupported alias cannot abort export.
    function assignEnumOptionSafely(options, propertyName, enumObject, candidates) {
        if (!enumObject) { return false; }
        var i;
        for (i = 0; i < candidates.length; i++) {
            var enumValue;
            try { enumValue = enumObject[candidates[i]]; } catch (ignoreEnumLookup) { enumValue = undefined; }
            if (enumValue === undefined || enumValue === null) { continue; }
            try {
                options[propertyName] = enumValue;
                return true;
            } catch (ignoreEnumAssignment) {
                // Try the next documented alias. If none work, retain Illustrator's safe default.
            }
        }
        return false;
    }

    function configurePDFSaveOptions(options, artboardRange) {
        // Use ordinary scalar properties directly; guard enum-typed properties separately.
        options.artboardRange = String(artboardRange || "");
        options.preserveEditability = false;
        options.generateThumbnails = true;
        options.viewAfterSaving = false;
        options.optimization = true;
        options.compressArt = true;

        var compatibilityEnum = null;
        var downsampleEnum = null;
        try { if (typeof PDFCompatibility !== "undefined") { compatibilityEnum = PDFCompatibility; } } catch (ignorePDFCompatibilityEnum) {}
        try { if (typeof DownsampleMethod !== "undefined") { downsampleEnum = DownsampleMethod; } } catch (ignoreDownsampleEnum) {}

        // Illustrator's actual enum members are usually uppercase (ACROBAT7). Some references
        // show title-case aliases, so try both and fall back to Acrobat 5/default if needed.
        assignEnumOptionSafely(options, "compatibility", compatibilityEnum, ["ACROBAT7", "Acrobat7", "ACROBAT5", "Acrobat5"]);

        // If a version does not expose a supported resampling enum, leave the host default rather
        // than failing the entire PDF export. Numeric downsampling values remain valid in that case.
        assignEnumOptionSafely(options, "colorDownsamplingMethod", downsampleEnum, ["BICUBICDOWNSAMPLE", "AVERAGEDOWNSAMPLE", "NODOWNSAMPLE"]);
        assignEnumOptionSafely(options, "grayscaleDownsamplingMethod", downsampleEnum, ["BICUBICDOWNSAMPLE", "AVERAGEDOWNSAMPLE", "NODOWNSAMPLE"]);
        assignEnumOptionSafely(options, "monochromeDownsamplingMethod", downsampleEnum, ["BICUBICDOWNSAMPLE", "AVERAGEDOWNSAMPLE", "NODOWNSAMPLE"]);

        options.colorDownsampling = 300;
        options.colorDownsamplingImageThreshold = 450;
        options.grayscaleDownsampling = 300;
        options.grayscaleDownsamplingImageThreshold = 450;
        options.monochromeDownsampling = 1200;
        options.monochromeDownsamplingImageThreshold = 1800;
        return options;
    }

    function exportOutputsAsSinglePDF(sourceDocument, items, sourceArtboardRect, sourceBounds) {
        if (!items || items.length < 1) { throw new Error("There are no outputs to export."); }
        if (items.length > MAX_ARTBOARDS) {
            throw new Error("A single multi-page PDF is limited to " + MAX_ARTBOARDS + " output pages by this Illustrator workflow. Use the separately numbered PDF export for larger batches.");
        }
        var pdfFile = File.saveDialog("Save all outputs into one multi-page PDF", "Adobe PDF:*.pdf");
        if (!pdfFile) { return { cancelled: true, count: 0 }; }
        if (!/\.pdf$/i.test(pdfFile.name)) { pdfFile = new File(pdfFile.fsName + ".pdf"); }
        var destinationExistedBefore = pdfFile.exists;
        if (destinationExistedBefore && !confirm("This PDF already exists. Overwrite it?", true, SCRIPT_NAME + " - PDF Export")) {
            return { cancelled: true, count: 0 };
        }

        var width = Number(sourceArtboardRect[2]) - Number(sourceArtboardRect[0]);
        var height = Number(sourceArtboardRect[1]) - Number(sourceArtboardRect[3]);
        if (!(width > 0) || !(height > 0)) { throw new Error("The source artboard dimensions are invalid."); }
        var currentDocument = app.activeDocument;
        var pdfDocument = null;
        var gap = 36; // Separate artboards in the temporary document; PDF page size remains exactly the source artboard size.
        try {
            pdfDocument = app.documents.add(sourceDocument.documentColorSpace, width, height);
            pdfDocument.artboards[0].artboardRect = [0, height, width, 0];
            var pdfColumns = Math.ceil(Math.sqrt(items.length));
            var i;
            for (i = 0; i < items.length; i++) {
                var boardColumn = i % pdfColumns;
                var boardRow = Math.floor(i / pdfColumns);
                var boardLeft = boardColumn * (width + gap);
                var boardTop = height - boardRow * (height + gap);
                var boardRect = [boardLeft, boardTop, boardLeft + width, boardTop - height];
                if (i > 0) { pdfDocument.artboards.add(boardRect); }
                else { pdfDocument.artboards[0].artboardRect = boardRect; }
                var duplicate = duplicatePageItemIntoDocument(items[i], sourceDocument, pdfDocument);
                fitPageItemWithinArtboard(duplicate, boardRect, 0.0);
                try { duplicate.name = "Output " + padNumber(i + 1, 3); } catch (ignoreCombinedPdfName) {}
            }

            var options = configurePDFSaveOptions(new PDFSaveOptions(), "1-" + items.length);
            pdfDocument.saveAs(pdfFile, options);
            pdfDocument.close(SaveOptions.DONOTSAVECHANGES);
            pdfDocument = null;
        } catch (singlePdfError) {
            if (pdfFile.exists && !destinationExistedBefore) {
                try { pdfFile.remove(); } catch (ignorePartialPdfCleanup) {}
            }
            throw new Error("Could not create the combined multi-page PDF. " + errorMessage(singlePdfError));
        } finally {
            try { if (pdfDocument) { pdfDocument.close(SaveOptions.DONOTSAVECHANGES); } } catch (ignoreCombinedPdfClose) {}
            try { app.activeDocument = sourceDocument; } catch (ignoreCombinedPdfRestore) {}
            if (!app.documents.length && currentDocument) { try { app.activeDocument = currentDocument; } catch (ignoreCombinedActiveRestore) {} }
        }
        return { file: pdfFile.fsName, count: items.length };
    }

    function exportRecordsAsSinglePDF(context, variableNames, recordRows, blankBehavior) {
        if (!context || !context.documentRef || !context.templateGroup) {
            throw new Error("The source artwork is not available. Return to Setup and select or restore the artwork Group.");
        }
        if (!recordRows || recordRows.length < 1) {
            throw new Error("Add or import at least one row before exporting a PDF.");
        }
        if (recordRows.length > MAX_ARTBOARDS) {
            throw new Error("A single PDF can contain up to " + MAX_ARTBOARDS + " pages in this Illustrator workflow. Export smaller batches for larger datasets.");
        }
        var sourceDocument = context.documentRef;
        var sourceItem = context.templateGroup;
        var sourceArtboardRect = context.sourceArtboardRect;
        var width = Number(sourceArtboardRect[2]) - Number(sourceArtboardRect[0]);
        var height = Number(sourceArtboardRect[1]) - Number(sourceArtboardRect[3]);
        if (!(width > 0) || !(height > 0)) { throw new Error("The source artboard has invalid dimensions."); }
        var pdfFile = File.saveDialog("Save all output items to one multi-page PDF", "PDF:*.pdf");
        if (!pdfFile) { return { cancelled: true, count: 0 }; }
        if (!/\.pdf$/i.test(pdfFile.name)) { pdfFile = new File(pdfFile.fsName + ".pdf"); }
        var existedBefore = pdfFile.exists;
        if (existedBefore && !confirm("This PDF already exists. Overwrite it?", true, SCRIPT_NAME + " - PDF Export")) {
            return { cancelled: true, count: 0 };
        }

        var currentDocument = app.activeDocument;
        var pdfDocument = null;
        var gap = 24; // Temporary spacing between PDF artboards; it does not change their page dimensions.
        var columns = Math.ceil(Math.sqrt(recordRows.length));
        var i;
        try {
            pdfDocument = app.documents.add(sourceDocument.documentColorSpace, width, height);
            pdfDocument.artboards[0].artboardRect = [0, height, width, 0];
            for (i = 0; i < recordRows.length; i++) {
                var column = i % columns;
                var row = Math.floor(i / columns);
                var left = column * (width + gap);
                var top = height - row * (height + gap);
                var pageRect = [left, top, left + width, top - height];
                if (i > 0) { pdfDocument.artboards.add(pageRect); }
                else { pdfDocument.artboards[0].artboardRect = pageRect; }

                var duplicate = duplicatePageItemIntoDocument(sourceItem, sourceDocument, pdfDocument);
                var mappedRecord = makeDataRowsFromGrid([recordRows[i]], variableNames).rows[0];
                replacePlaceholders(duplicate, variableNames, context.config.frameIndices, mappedRecord, blankBehavior || "Clear mapped text", 1);
                fitPageItemWithinArtboard(duplicate, pageRect, 0.0);
                try { duplicate.name = "Output_" + padNumber(i + 1, 3); } catch (ignorePDFOutputName) {}
            }

            var options = configurePDFSaveOptions(new PDFSaveOptions(), "1-" + recordRows.length);
            pdfDocument.saveAs(pdfFile, options);
            pdfDocument.close(SaveOptions.DONOTSAVECHANGES);
            pdfDocument = null;
        } catch (exportError) {
            if (pdfFile.exists && !existedBefore) {
                try { pdfFile.remove(); } catch (ignorePartialPdfDelete) {}
            }
            throw new Error("PDF export failed. The data window remains open. " + errorMessage(exportError));
        } finally {
            try { if (pdfDocument) { pdfDocument.close(SaveOptions.DONOTSAVECHANGES); } } catch (ignorePDFTempClose) {}
            try { app.activeDocument = sourceDocument; } catch (ignorePDFSourceRestore) {}
            if (!app.documents.length && currentDocument) { try { app.activeDocument = currentDocument; } catch (ignorePDFCurrentRestore) {} }
        }
        return { file: pdfFile.fsName, count: recordRows.length };
    }

    function exportOutputsAsNumberedPDFs(sourceDocument, items, sourceArtboardRect, sourceBounds) {
        var folder = Folder.selectDialog("Choose the folder for numbered output PDF files");
        if (!folder) { return { count: 0, cancelled: true }; }
        var requestedPrefix = prompt("Filename prefix for the PDF files:", "Output", SCRIPT_NAME + " - PDF Export");
        if (requestedPrefix === null || trim(String(requestedPrefix)) === "") { return { count: 0, cancelled: true }; }
        var prefix = sanitizePresetFileName(requestedPrefix);
        var files;
        function buildFileList(namePrefix) {
            var list = [];
            var j;
            for (j = 0; j < items.length; j++) {
                list.push(new File(folder.fsName + "/" + namePrefix + "_" + padNumber(j + 1, 3) + ".pdf"));
            }
            return list;
        }
        function hasCollisions(list) {
            var j;
            for (j = 0; j < list.length; j++) { if (list[j].exists) { return true; } }
            return false;
        }
        files = buildFileList(prefix);
        if (hasCollisions(files)) {
            if (confirm("Some numbered PDF files already exist. Overwrite matching files?", true, SCRIPT_NAME + " - PDF Export")) {
                // Intentional overwrite after explicit user confirmation.
            } else {
                var suffix = 2;
                var originalPrefix = prefix;
                do {
                    prefix = originalPrefix + "_" + suffix;
                    files = buildFileList(prefix);
                    suffix++;
                } while (hasCollisions(files) && suffix < 10000);
                if (hasCollisions(files)) { throw new Error("Could not find an unused output filename prefix."); }
            }
        }

        var boardWidth = sourceArtboardRect[2] - sourceArtboardRect[0];
        var boardHeight = sourceArtboardRect[1] - sourceArtboardRect[3];
        var exportedCount = 0;
        var i;
        for (i = 0; i < items.length; i++) {
            var temporaryDocument = null;
            try {
                temporaryDocument = makeTemporaryOutputDocument(
                    sourceDocument, items[i], sourceArtboardRect, sourceBounds, null, null, null, "Clear mapped text"
                );
                // The temporary document has exactly one artboard whose dimensions match the source output.
                var options = configurePDFSaveOptions(new PDFSaveOptions(), "1");
                temporaryDocument.saveAs(files[i], options);
                temporaryDocument.close(SaveOptions.DONOTSAVECHANGES);
                temporaryDocument = null;
                exportedCount++;
            } catch (pdfOneError) {
                try { if (temporaryDocument) { temporaryDocument.close(SaveOptions.DONOTSAVECHANGES); } } catch (ignorePdfClose) {}
                try { app.activeDocument = sourceDocument; } catch (ignorePdfRestore) {}
                throw new Error("PDF export stopped at output " + (i + 1) + " of " + items.length + ". " +
                    exportedCount + " file(s) were already written. Details: " + errorMessage(pdfOneError));
            }
        }
        try { app.activeDocument = sourceDocument; } catch (ignoreFinalPDFRestore) {}
        return { count: exportedCount, folder: folder.fsName, prefix: prefix };
    }

    // Run the Excel COM helper through cscript explicitly instead of relying on the
    // Windows file association for .vbs files. This also captures compiler/runtime errors.
    function runWindowsVbsHelper(helperFile, statusFile, logFile, description) {
        var launcherFile = new File(helperFile.fsName.replace(/\.vbs$/i, ".cmd"));
        try { if (statusFile.exists) { statusFile.remove(); } } catch (ignoreOldStatus) {}
        try { if (logFile.exists) { logFile.remove(); } } catch (ignoreOldLog) {}
        var commandText = "@echo off\r\n" +
            '"%SystemRoot%\\System32\\cscript.exe" //nologo //B "%~dp0' + helperFile.name + '" > "%~dp0' + logFile.name + '" 2>&1\r\n' +
            "exit /b %errorlevel%\r\n";
        launcherFile.encoding = "ASCII";
        if (!launcherFile.open("w")) { throw new Error("Could not create the Windows Excel launcher."); }
        launcherFile.write(commandText);
        launcherFile.close();

        var launched = false;
        var launchedDirectVbs = false;
        try { launched = launcherFile.execute(); } catch (launchError) { launched = false; }
        if (!launched) {
            // Some Windows configurations do not associate .cmd with ShellExecute correctly.
            // Fall back to the Windows Script Host association for the .vbs file itself.
            try { launched = helperFile.execute(); launchedDirectVbs = launched; } catch (fallbackLaunchError) { launched = false; }
        }
        if (launched && !launchedDirectVbs && !statusFile.exists && !logFile.exists) {
            // File.execute() can return true even if the .cmd association only opens a window.
            // Give the launcher a brief start window, then try the .vbs association if nothing ran.
            var launchProbeStart = new Date().getTime();
            while (!statusFile.exists && !logFile.exists && new Date().getTime() - launchProbeStart < 5000) { $.sleep(250); }
            if (!statusFile.exists && !logFile.exists) {
                try { launched = helperFile.execute(); launchedDirectVbs = launched; } catch (secondFallbackLaunchError) { launched = false; }
            }
        }
        if (!launched) {
            try { launcherFile.remove(); } catch (ignoreLauncherLaunchFailure) {}
            throw new Error("Windows could not start the Excel import helper. Check Windows Script Host and Microsoft Excel availability.");
        }
        var startedAt = new Date().getTime();
        var timeoutMs = 90000;
        while (!statusFile.exists && new Date().getTime() - startedAt < timeoutMs) { $.sleep(250); }
        if (!statusFile.exists) {
            var diagnostic = "";
            try {
                if (logFile.exists) {
                    logFile.encoding = "UTF-8";
                    if (logFile.open("r")) { diagnostic = logFile.read(); logFile.close(); }
                }
            } catch (ignoreLogRead) {}
            try { launcherFile.remove(); } catch (ignoreLauncherTimeout) {}
            throw new Error("Excel import timed out after 90 seconds. Confirm Microsoft Excel is installed and no modal Excel dialog is open." +
                (diagnostic ? "\n\nWindows script details:\n" + diagnostic.substring(0, 1800) : "") +
                "\n\nA fallback is to save the worksheet as CSV UTF-8 or Tab Delimited Text.");
        }
        var statusText = "";
        statusFile.encoding = "UTF-8";
        if (statusFile.open("r")) { statusText = statusFile.read(); statusFile.close(); }
        statusText = String(statusText).replace(/^\uFEFF/, "");
        var logText = "";
        try {
            if (logFile.exists) {
                logFile.encoding = "UTF-8";
                if (logFile.open("r")) { logText = logFile.read(); logFile.close(); }
            }
        } catch (ignoreLogReadAfterRun) {}
        try { launcherFile.remove(); } catch (ignoreLauncherCleanup) {}
        if (statusText.indexOf("ERROR|") === 0) {
            try { if (launcherFile.exists) { launcherFile.remove(); } } catch (ignoreLauncherErrorCleanup) {}
            throw new Error(statusText.substring(6) + (logText ? "\n\nWindows script details:\n" + logText.substring(0, 1800) : ""));
        }
        if (statusText.indexOf("OK") !== 0) {
            throw new Error("The Excel helper returned an unexpected status." + (logText ? "\n\n" + logText.substring(0, 1800) : ""));
        }
        try { if (logFile.exists) { logFile.remove(); } } catch (ignoreLogCleanup) {}
        return statusText;
    }

    // --------------------------- Two-stage UI ---------------------------
    function runTwoStageWizard(textFrames, groupName, documentRef, templateGroup, sourceArtboardRect, sourceBounds, initialPreset) {
        var activePreset = initialPreset || null;
        var lastRecords = null;
        var lastBlankBehavior = "Clear mapped text";
        var loopGuard = 0;

        while (loopGuard++ < 20) {
            var setupWindow = buildSetupDialog(textFrames, groupName, activePreset);
            if (setupWindow.show() !== 1) { return null; }

            var config;
            try {
                config = readConfiguration(setupWindow);
            } catch (setupReadError) {
                alert(errorMessage(setupReadError), SCRIPT_NAME + " - Setup");
                return null;
            }

            var loadedPreset = setupWindow.getLoadedPreset();
            var dataContext = {
                config: config,
                frameIndices: config.frameIndices,
                templateFrames: textFrames,
                documentRef: documentRef,
                templateGroup: templateGroup,
                groupName: groupName,
                sourceArtboardRect: copyRect(sourceArtboardRect),
                sourceBounds: copyRect(sourceBounds)
            };
            var startingRecords = loadedPreset ? loadedPreset.records : lastRecords;
            var startingBlank = loadedPreset ? loadedPreset.blankBehavior : lastBlankBehavior;
            var dataWindow = buildDataDialog(
                config.variables,
                startingRecords,
                startingBlank || "Clear mapped text",
                loadedPreset ? loadedPreset.columnWidths : null,
                dataContext
            );
            var dataStatus = dataWindow.show();

            if (dataStatus === 2) { return { saveOnly: true }; }
            if (dataStatus === 3) {
                // Preserve all edited records and the current mappings/settings before reopening Stage 1.
                lastRecords = dataWindow.getRecords();
                lastBlankBehavior = dataWindow.getBlankBehavior();
                activePreset = makeTransientWizardPreset(
                    config, textFrames, groupName, sourceArtboardRect,
                    lastRecords, lastBlankBehavior, loadedPreset
                );
                continue;
            }
            if (dataStatus !== 1) { return null; }

            return {
                config: config,
                records: dataWindow.getRecords(),
                blankBehavior: dataWindow.getBlankBehavior(),
                pdfMode: dataWindow.getPdfMode ? dataWindow.getPdfMode() : "prompt"
            };
        }
        alert("The setup was reopened too many times. Please run the script again.", SCRIPT_NAME);
        return null;
    }

    function makeTransientWizardPreset(config, textFrames, groupName, sourceArtboardRect, records, blankBehavior, previousPreset) {
        var maps = [];
        var i;
        for (i = 0; i < config.variables.length; i++) {
            var frameIndex = config.frameIndices[i];
            var frame = (frameIndex >= 0 && frameIndex < textFrames.length) ? textFrames[frameIndex] : null;
            var metadata = frame ? getTextFrameMetadata(frame, frameIndex) : {
                index: frameIndex, contents: "", layerName: "", bounds: ""
            };
            maps.push({
                variableName: String(config.variables[i]),
                frameIndex: frameIndex,
                sourceContents: metadata.contents,
                layerName: metadata.layerName,
                bounds: metadata.bounds,
                normalizedBounds: frame ? normalizeBoundsToArtboard(metadata.bounds, sourceArtboardRect) : ""
            });
        }
        return {
            name: previousPreset && previousPreset.name ? previousPreset.name : (groupName || "Working Setup"),
            created: new Date().toString(),
            templateName: String(groupName || ""),
            templateBackupPath: previousPreset ? String(previousPreset.templateBackupPath || "") : "",
            blankBehavior: String(blankBehavior || "Clear mapped text"),
            config: config,
            variables: config.variables.slice(0),
            maps: maps,
            records: cloneRecordRows(records || [], config.variables.length),
            columnWidths: [],
            sourceArtboardRect: copyRect(sourceArtboardRect),
            sourceDocumentName: safeDocumentName(app.activeDocument)
        };
    }

    function buildSetupDialog(textFrames, groupName, initialPreset) {
        var win = new Window("dialog", SCRIPT_NAME + " " + SCRIPT_VERSION + " - 1/2: Setup");
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.spacing = 4;
        win.margins = 9;
        win.preferredSize = [760, 520];

        var introRow = win.add("group");
        introRow.orientation = "row";
        introRow.alignChildren = ["left", "center"];
        introRow.spacing = 8;
        var intro = introRow.add("statictext", undefined,
            "Template: " + (groupName ? groupName : "(unnamed)") +
            " | Text frames: " + textFrames.length,
            { multiline: false }
        );
        intro.preferredSize.height = 18;
        intro.alignment = ["fill", "center"];
        var loadPresetButton = introRow.add("button", undefined, "Load Preset...");
        loadPresetButton.preferredSize.width = 105;

        var mappingPanel = win.add("panel", undefined, "1. Variable Mapping");
        mappingPanel.orientation = "column";
        mappingPanel.alignChildren = ["fill", "top"];
        mappingPanel.margins = 8;
        mappingPanel.spacing = 4;

        var mappingHeader = mappingPanel.add("group");
        mappingHeader.orientation = "row";
        mappingHeader.alignChildren = ["center", "center"];
        mappingHeader.add("statictext", undefined, "Variable name").preferredSize.width = 125;
        mappingHeader.add("statictext", undefined, "Select a text frame").preferredSize.width = 480;

        var mappingViewport = mappingPanel.add("group");
        mappingViewport.orientation = "row";
        mappingViewport.alignChildren = ["left", "top"];
        mappingViewport.spacing = 4;
        // Fixed-height viewport: five complete mapping rows remain visible.
        mappingViewport.preferredSize = [638, 154];
        mappingViewport.maximumSize = [700, 154];

        var mappingRowsGroup = mappingViewport.add("group");
        mappingRowsGroup.orientation = "column";
        mappingRowsGroup.alignChildren = ["fill", "top"];
        mappingRowsGroup.spacing = 3;
        mappingRowsGroup.preferredSize = [610, 150];
        mappingRowsGroup.maximumSize = [680, 150];

        var mappingVerticalScrollbar = mappingViewport.add("scrollbar", [0, 0, 16, 150], 0, 0, 0);
        mappingVerticalScrollbar.preferredSize = [16, 150];
        mappingVerticalScrollbar.minimumSize = [16, 150];
        mappingVerticalScrollbar.maximumSize = [16, 150];
        mappingVerticalScrollbar.helpTip = "Scroll through variable mappings.";

        var mappingNavigation = mappingPanel.add("group");
        mappingNavigation.orientation = "row";
        mappingNavigation.alignChildren = ["center", "center"];
        mappingNavigation.spacing = 6;
        var previousMappingButton = mappingNavigation.add("button", undefined, "< Previous");
        var mappingPageLabel = mappingNavigation.add("statictext", undefined, "Variables 1-1");
        var nextMappingButton = mappingNavigation.add("button", undefined, "Next >");

        var mappingButtons = mappingPanel.add("group");
        mappingButtons.orientation = "row";
        mappingButtons.alignChildren = ["left", "center"];
        var addVariableButton = mappingButtons.add("button", undefined, "Add Variable");
        var removeVariableButton = mappingButtons.add("button", undefined, "Delete Variables...");

        var optionLabels = [];
        var i;
        for (i = 0; i < textFrames.length; i++) {
            optionLabels.push(makeTextFrameLabel(textFrames[i], i));
        }
        // mappingRows is the persistent model, not the visual control list.
        // Only the five rows inside the viewport are instantiated as ScriptUI controls.
        var mappingRows = [];
        var mappingSlots = [];
        var mappingScrollOffset = 0;
        var mappingRowsPerPage = 5;
        var loadedPreset = null;

        var layoutPanel = win.add("panel", undefined, "2. Output Layout");
        layoutPanel.orientation = "column";
        layoutPanel.alignChildren = ["fill", "top"];
        layoutPanel.margins = 8;
        layoutPanel.spacing = 5;

        var layoutRow = layoutPanel.add("group");
        layoutRow.orientation = "row";
        layoutRow.alignChildren = ["center", "center"];
        layoutRow.add("statictext", undefined, "Layout:");
        var layoutDropdown = layoutRow.add("dropdownlist", undefined,
            ["Separate Artboards", "Grid", "Horizontal", "Vertical", "Same Position"]
        );
        layoutDropdown.selection = 1;
        layoutDropdown.preferredSize.width = 205;

        var artboardArrangementRow = layoutPanel.add("group");
        artboardArrangementRow.orientation = "row";
        artboardArrangementRow.alignChildren = ["center", "center"];
        artboardArrangementRow.add("statictext", undefined, "Artboard arrangement:");
        var artboardArrangementDropdown = artboardArrangementRow.add("dropdownlist", undefined, ["Grid", "Horizontal only", "Vertical only"]);
        artboardArrangementDropdown.selection = 0;
        artboardArrangementDropdown.preferredSize.width = 165;

        var gridGroup = layoutPanel.add("group");
        gridGroup.orientation = "row";
        gridGroup.alignChildren = ["center", "center"];
        var columnsInput = addLabeledInput(gridGroup, "Columns:", "5", 4);
        var rowsInput = addLabeledInput(gridGroup, "Rows:", "8", 4);
        var horizontalGridSpacingInput = addLabeledInput(gridGroup, "Horizontal gap (cm):", "1", 5);
        var verticalGridSpacingInput = addLabeledInput(gridGroup, "Vertical gap (cm):", "1", 5);

        var horizontalGroup = layoutPanel.add("group");
        horizontalGroup.orientation = "row";
        horizontalGroup.alignChildren = ["center", "center"];
        var horizontalSpacingInput = addLabeledInput(horizontalGroup, "Space between outputs (cm):", "1", 5);

        var verticalGroup = layoutPanel.add("group");
        verticalGroup.orientation = "row";
        verticalGroup.alignChildren = ["center", "center"];
        var verticalSpacingInput = addLabeledInput(verticalGroup, "Space between outputs (cm):", "1", 5);

        var layoutHelp = layoutPanel.add("statictext", undefined, "", { multiline: false });
        layoutHelp.preferredSize.height = 16;

        var templatePanel = win.add("panel", undefined, "3. Output and Template");
        templatePanel.orientation = "column";
        templatePanel.alignChildren = ["left", "top"];
        templatePanel.margins = 6;
        templatePanel.spacing = 3;

        var performanceRow = templatePanel.add("group");
        performanceRow.orientation = "row";
        performanceRow.alignChildren = ["center", "center"];
        var flattenStaticCheckbox = performanceRow.add("checkbox", undefined,
            "Flatten static artwork; keep mapped text editable");
        flattenStaticCheckbox.value = false;
        performanceRow.add("statictext", undefined, "Resolution:");
        var rasterResolutionDropdown = performanceRow.add("dropdownlist", undefined, ["150 ppi", "300 ppi", "600 ppi"]);
        rasterResolutionDropdown.selection = 1;
        rasterResolutionDropdown.preferredSize.width = 90;
        rasterResolutionDropdown.enabled = false;
        flattenStaticCheckbox.onClick = function () {
            rasterResolutionDropdown.enabled = flattenStaticCheckbox.value;
            refreshActionHelp();
        };
        var flattenHelp = templatePanel.add("statictext", undefined,
            "Optional: rasterizes non-variable artwork; mapped text remains editable.",
            { multiline: false }
        );
        // Give the helper enough width so the full sentence remains on one line.
        // A fixed 18px height previously clipped the wrapped second line in ScriptUI.
        flattenHelp.preferredSize.width = 650;
        flattenHelp.preferredSize.height = 22;

        var actionRow = templatePanel.add("group");
        actionRow.orientation = "row";
        actionRow.alignChildren = ["center", "center"];
        actionRow.add("statictext", undefined, "After generation:");
        var actionDropdown = actionRow.add("dropdownlist", undefined,
            ["Keep Original", "Delete Original", "Replace Original"]
        );
        actionDropdown.selection = 0;
        actionDropdown.preferredSize.width = 190;
        var actionHelp = templatePanel.add("statictext", undefined, "", { multiline: false });
        actionHelp.preferredSize.height = 16;

        var buttonRow = win.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["right", "top"];
        buttonRow.spacing = 8;
        var cancelButton = buttonRow.add("button", undefined, "Cancel", { name: "cancel" });
        var continueButton = buttonRow.add("button", undefined, "Continue to Data >", { name: "ok" });

        function syncMappingRowsFromControls() {
            // Copy the five active UI slots into their corresponding model entries.
            // Off-screen variables are never represented by hidden ScriptUI controls.
            var i;
            for (i = 0; i < mappingSlots.length; i++) {
                var slot = mappingSlots[i];
                if (slot.modelIndex >= 0 && slot.modelIndex < mappingRows.length) {
                    var model = mappingRows[slot.modelIndex];
                    model.name = String(slot.nameInput.text);
                    model.frameIndex = (slot.dropdown.selection && slot.dropdown.selection.index > 0)
                        ? slot.dropdown.selection.index - 1 : -1;
                }
            }
        }

        function createMappingRow(slotIndex) {
            // Create a permanent visual slot. The slot is rebound to a model item
            // while scrolling; its ScriptUI controls are never removed/re-added.
            var row = mappingRowsGroup.add("group");
            row.orientation = "row";
            row.alignChildren = ["center", "center"];
            row.alignment = ["left", "top"];
            row.spacing = 6;
            row.preferredSize = [610, 27];
            row.minimumSize = [610, 27];
            row.maximumSize = [680, 27];

            var nameInput = row.add("edittext", undefined, "");
            nameInput.characters = 17;
            nameInput.preferredSize = [125, 25];
            nameInput.minimumSize = [125, 25];
            nameInput.maximumSize = [125, 25];

            var dropdown = row.add("dropdownlist", undefined, ["< Select a text frame >"].concat(optionLabels));
            dropdown.selection = 0;
            dropdown.preferredSize = [470, 25];
            dropdown.minimumSize = [470, 25];
            dropdown.maximumSize = [470, 25];

            var slot = {
                slotIndex: slotIndex,
                modelIndex: -1,
                group: row,
                nameInput: nameInput,
                dropdown: dropdown
            };
            (function (boundSlot) {
                boundSlot.nameInput.onChange = function () {
                    if (boundSlot.modelIndex >= 0 && mappingRows[boundSlot.modelIndex]) {
                        mappingRows[boundSlot.modelIndex].name = String(boundSlot.nameInput.text);
                    }
                };
                boundSlot.dropdown.onChange = function () {
                    if (boundSlot.modelIndex >= 0 && mappingRows[boundSlot.modelIndex]) {
                        mappingRows[boundSlot.modelIndex].frameIndex =
                            (boundSlot.dropdown.selection && boundSlot.dropdown.selection.index > 0)
                            ? boundSlot.dropdown.selection.index - 1 : -1;
                    }
                };
            }(slot));
            mappingSlots.push(slot);
        }

        function refreshMappingPage(skipSync) {
            if (!skipSync) { syncMappingRowsFromControls(); }
            var maxOffset = Math.max(0, mappingRows.length - mappingRowsPerPage);
            if (mappingScrollOffset > maxOffset) { mappingScrollOffset = maxOffset; }
            if (mappingScrollOffset < 0) { mappingScrollOffset = 0; }
            var start = mappingScrollOffset;
            var end = Math.min(mappingRows.length, start + mappingRowsPerPage);

            var slotIndex;
            for (slotIndex = 0; slotIndex < mappingSlots.length; slotIndex++) {
                var slot = mappingSlots[slotIndex];
                var modelIndex = start + slotIndex;
                if (modelIndex < mappingRows.length && slotIndex < mappingRowsPerPage) {
                    var model = mappingRows[modelIndex];
                    slot.modelIndex = modelIndex;
                    slot.group.visible = true;
                    slot.nameInput.enabled = true;
                    slot.dropdown.enabled = true;
                    slot.nameInput.text = String(model.name);
                    slot.dropdown.selection = (model.frameIndex !== undefined && model.frameIndex >= 0 && model.frameIndex < textFrames.length)
                        ? model.frameIndex + 1 : 0;
                } else {
                    slot.modelIndex = -1;
                    slot.nameInput.text = "";
                    slot.dropdown.selection = 0;
                    slot.nameInput.enabled = false;
                    slot.dropdown.enabled = false;
                    // Unused slots are only ever at the bottom of the fixed viewport,
                    // so they cannot create gaps before a later visible row.
                    slot.group.visible = false;
                }
            }

            mappingPageLabel.text = "Variables " + (mappingRows.length ? (start + 1) : 0) + "-" + end + " of " + mappingRows.length;
            previousMappingButton.enabled = mappingScrollOffset > 0;
            nextMappingButton.enabled = mappingScrollOffset < maxOffset;
            mappingVerticalScrollbar.minvalue = 0;
            mappingVerticalScrollbar.maxvalue = maxOffset;
            mappingVerticalScrollbar.stepdelta = 1;
            mappingVerticalScrollbar.jumpdelta = 1;
            mappingVerticalScrollbar.value = mappingScrollOffset;
            mappingVerticalScrollbar.enabled = maxOffset > 0;
            removeVariableButton.enabled = mappingRows.length > 1;
            try { win.update(); } catch (ignoreMappingViewportUpdate) {}
        }

        function createMappingModel(defaultName, selectedFrameIndex) {
            mappingRows.push({
                name: String(defaultName || ""),
                frameIndex: selectedFrameIndex !== undefined ? Number(selectedFrameIndex) : -1
            });
        }

        function addVariable() {
            if (mappingRows.length >= textFrames.length) {
                alert("You cannot add more variables than there are text frames in the selected Group.", SCRIPT_NAME);
                return;
            }
            syncMappingRowsFromControls();
            createMappingModel("V" + (mappingRows.length + 1), -1);
            mappingScrollOffset = Math.max(0, mappingRows.length - mappingRowsPerPage);
            refreshMappingPage(true);
        }

        function openDeleteVariablesDialog() {
            syncMappingRowsFromControls();
            if (mappingRows.length <= 1) {
                alert("At least one variable must remain. You can delete one or more variables, but not all of them.", SCRIPT_NAME + " - Delete Variables");
                return;
            }
            var deleteWin = new Window("dialog", SCRIPT_NAME + " - Delete Variables");
            deleteWin.orientation = "column";
            deleteWin.alignChildren = ["fill", "top"];
            deleteWin.spacing = 7;
            deleteWin.margins = 11;
            deleteWin.preferredSize = [390, 120];
            deleteWin.add("statictext", undefined, "Variables to delete (current serial numbers):");
            var specInput = deleteWin.add("edittext", undefined, "");
            specInput.characters = 34;
            specInput.active = true;
            deleteWin.add("statictext", undefined, "Use 2, 4 or ranges such as 2-5,7. The remaining order is preserved.");
            var buttons = deleteWin.add("group");
            buttons.orientation = "row";
            buttons.alignment = ["right", "top"];
            var cancel = buttons.add("button", undefined, "Cancel", { name: "cancel" });
            var ok = buttons.add("button", undefined, "Delete", { name: "ok" });
            var indexes = null;
            cancel.onClick = function () { deleteWin.close(0); };
            ok.onClick = function () {
                try {
                    indexes = parseIndexSpecification(specInput.text, mappingRows.length, "Variable");
                    if (!indexes.length) { throw new Error("Enter at least one variable number."); }
                    if (indexes.length >= mappingRows.length) { throw new Error("At least one variable must remain."); }
                    deleteWin.close(1);
                } catch (err) { alert(errorMessage(err), SCRIPT_NAME + " - Delete Variables"); }
            };
            if (deleteWin.show() !== 1 || !indexes) { return; }
            if (!confirm("Delete " + indexes.length + " variable mapping(s)? Any saved data columns for these variables will also be removed.", true, SCRIPT_NAME)) { return; }

            var deleteMap = {};
            var d;
            for (d = 0; d < indexes.length; d++) { deleteMap[String(indexes[d])] = true; }
            var kept = [];
            var oldIndex;
            for (oldIndex = 0; oldIndex < mappingRows.length; oldIndex++) {
                if (!deleteMap[String(oldIndex + 1)]) {
                    kept.push({
                        name: trim(String(mappingRows[oldIndex].name)),
                        frameIndex: Number(mappingRows[oldIndex].frameIndex),
                        oldIndex: oldIndex
                    });
                }
            }

            // Keep preset records aligned with the surviving variable columns.
            if (loadedPreset) {
                var nextVariables = [];
                var nextMaps = [];
                var nextWidths = [];
                var nextRecords = [];
                var k;
                for (k = 0; k < kept.length; k++) {
                    var priorIndex = kept[k].oldIndex;
                    nextVariables.push(kept[k].name);
                    var savedMap = loadedPreset.maps[priorIndex] || {
                        variableName: kept[k].name, frameIndex: kept[k].frameIndex, sourceContents: "", layerName: "", bounds: "", normalizedBounds: ""
                    };
                    savedMap.variableName = kept[k].name;
                    savedMap.frameIndex = kept[k].frameIndex;
                    nextMaps.push(savedMap);
                    nextWidths.push(loadedPreset.columnWidths && loadedPreset.columnWidths[priorIndex] !== undefined ? loadedPreset.columnWidths[priorIndex] : 0);
                }
                for (k = 0; k < loadedPreset.records.length; k++) {
                    var nextRow = [];
                    var c;
                    for (c = 0; c < kept.length; c++) {
                        var oldCellIndex = kept[c].oldIndex;
                        nextRow.push(loadedPreset.records[k] && loadedPreset.records[k][oldCellIndex] !== undefined ? String(loadedPreset.records[k][oldCellIndex]) : "");
                    }
                    nextRecords.push(nextRow);
                }
                loadedPreset.variables = nextVariables;
                loadedPreset.maps = nextMaps;
                loadedPreset.columnWidths = nextWidths;
                loadedPreset.records = nextRecords;
            }

            mappingRows = [];
            for (oldIndex = 0; oldIndex < kept.length; oldIndex++) {
                createMappingModel(kept[oldIndex].name, kept[oldIndex].frameIndex);
            }
            mappingScrollOffset = Math.min(mappingScrollOffset, Math.max(0, kept.length - mappingRowsPerPage));
            refreshMappingPage(true);
        }

        function readMappings() {
            var variables = [];
            var frameIndices = [];
            var usedIndices = [];
            var usedNames = [];
            syncMappingRowsFromControls();
            var m;
            for (m = 0; m < mappingRows.length; m++) {
                var variableName = trim(String(mappingRows[m].name));
                if (!variableName) {
                    throw new Error("Variable " + (m + 1) + " needs a name.");
                }
                if (/[\t\r\n]/.test(variableName)) {
                    throw new Error("Variable names cannot contain tabs or line breaks.");
                }
                if (/^(?:__proto__|constructor|prototype|tostring|valueof|hasownproperty|isprototypeof|propertyisenumerable|tolocalestring)$/i.test(variableName)) {
                    throw new Error("Please choose a different name for '" + variableName + "'. This name is reserved by JavaScript.");
                }
                var normalizedName = variableName.toLowerCase().replace(/\s+/g, " ");
                if (contains(usedNames, normalizedName)) {
                    throw new Error("Variable names must be unique. Duplicate name: " + variableName);
                }
                usedNames.push(normalizedName);

                var frameIndex = Number(mappingRows[m].frameIndex);
                if (isNaN(frameIndex) || frameIndex < 0 || frameIndex >= textFrames.length) {
                    throw new Error("Please assign '" + variableName + "' to a text frame before continuing.");
                }
                if (contains(usedIndices, frameIndex)) {
                    throw new Error("Each variable must use a different text frame. The same text frame was assigned more than once.");
                }
                usedIndices.push(frameIndex);
                variables.push(variableName);
                frameIndices.push(frameIndex);
            }
            return { variables: variables, frameIndices: frameIndices };
        }

        function applyPresetToControls(preset) {
            var i;
            mappingRows = [];
            mappingScrollOffset = 0;
            for (i = 0; i < preset.variables.length; i++) {
                var mappedIndex = findPresetFrameMatch(preset.maps[i], textFrames);
                createMappingModel(preset.variables[i], mappedIndex);
            }
            if (mappingRows.length === 0) { createMappingModel("V1", -1); }
            mappingScrollOffset = 0;
            refreshMappingPage(true);
            removeVariableButton.enabled = (mappingRows.length > 1);

            setDropdownText(layoutDropdown, preset.config.layout);
            setDropdownText(artboardArrangementDropdown, preset.config.artboardArrangement || "Grid");
            setDropdownText(actionDropdown, preset.config.action);
            columnsInput.text = String(preset.config.columns || 5);
            rowsInput.text = String(preset.config.rows || 8);
            horizontalGridSpacingInput.text = String(preset.config.gridGapX === undefined ? 1 : preset.config.gridGapX);
            verticalGridSpacingInput.text = String(preset.config.gridGapY === undefined ? 1 : preset.config.gridGapY);
            horizontalSpacingInput.text = String(preset.config.horizontalGap === undefined ? 1 : preset.config.horizontalGap);
            verticalSpacingInput.text = String(preset.config.verticalGap === undefined ? 1 : preset.config.verticalGap);
            flattenStaticCheckbox.value = Boolean(preset.config.flattenStaticArtwork);
            rasterResolutionDropdown.enabled = flattenStaticCheckbox.value;
            setDropdownText(rasterResolutionDropdown, String(preset.config.rasterResolution || 300) + " ppi");
            loadedPreset = preset;
            refreshLayoutControls();
            refreshActionHelp();
        }

        function setDropdownText(dropdown, text) {
            var i;
            for (i = 0; i < dropdown.items.length; i++) {
                if (String(dropdown.items[i].text) === String(text)) {
                    dropdown.selection = i;
                    return;
                }
            }
        }

        loadPresetButton.onClick = function () {
            var presetFile = choosePresetFile("Load a Variable Data Studio preset (.vds or legacy .cbg)");
            if (!presetFile) { return; }
            try {
                var preset = readPresetFile(presetFile);
                rememberLastPresetPath(presetFile);
                applyPresetToControls(preset);
                var mappedCount = 0;
                for (var p = 0; p < preset.maps.length; p++) {
                    if (findPresetFrameMatch(preset.maps[p], textFrames) >= 0) { mappedCount++; }
                }
                var mappingMessage = mappedCount + " of " + preset.variables.length + " text mapping(s) matched this template.";
                if (mappedCount < preset.variables.length) {
                    mappingMessage += "\nPlease reassign the unmatched variables manually before continuing.";
                }
                if (preset.templateBackupPath) {
                    mappingMessage += "\nTemplate backup: " + preset.templateBackupPath;
                }
                alert("Preset loaded.\n" + mappingMessage, SCRIPT_NAME + " - Load Preset");
            } catch (loadError) {
                alert(errorMessage(loadError), SCRIPT_NAME + " - Load Preset");
            }
        };

        function refreshLayoutControls() {
            var mode = layoutDropdown.selection ? layoutDropdown.selection.text : "Grid";
            var artboardArrangement = artboardArrangementDropdown.selection ? artboardArrangementDropdown.selection.text : "Grid";
            artboardArrangementRow.visible = (mode === "Separate Artboards");
            gridGroup.visible = (mode === "Grid" || mode === "Separate Artboards");
            var showGridDimensions = (mode === "Grid") || (mode === "Separate Artboards" && artboardArrangement === "Grid");
            columnsInput.parent.visible = showGridDimensions;
            rowsInput.parent.visible = showGridDimensions;
            horizontalGridSpacingInput.parent.enabled = !(mode === "Separate Artboards" && artboardArrangement === "Vertical only");
            verticalGridSpacingInput.parent.enabled = !(mode === "Separate Artboards" && artboardArrangement === "Horizontal only");
            horizontalGroup.visible = (mode === "Horizontal");
            verticalGroup.visible = (mode === "Vertical");
            if (mode === "Separate Artboards") {
                if (artboardArrangement === "Horizontal only") { layoutHelp.text = "One artboard per record, arranged in a single horizontal row."; }
                else if (artboardArrangement === "Vertical only") { layoutHelp.text = "One artboard per record, arranged in a single vertical column."; }
                else { layoutHelp.text = "One artboard per record; set columns, rows, and both gaps above."; }
            } else if (mode === "Grid") {
                layoutHelp.text = "Set columns, rows, and horizontal/vertical gaps.";
            } else if (mode === "Horizontal") {
                layoutHelp.text = "One horizontal row; set the gap above.";
            } else if (mode === "Vertical") {
                layoutHelp.text = "One vertical column; set the gap above.";
            } else {
                layoutHelp.text = "All copies overlap at the same position.";
            }
            refreshActionHelp();
            win.layout.layout(true);
        }

        function refreshActionHelp() {
            var mode = layoutDropdown.selection ? layoutDropdown.selection.text : "Grid";
            var action = actionDropdown.selection ? actionDropdown.selection.text : "Keep Original";
            if (action === "Keep Original") {
                actionHelp.text = "Keep the selected template and create copies.";
            } else if (action === "Delete Original") {
                actionHelp.text = "Delete the selected template after success.";
            } else if (mode === "Separate Artboards" && flattenStaticCheckbox.value) {
                actionHelp.text = "Replace the template with the first flattened record.";
            } else if (mode === "Separate Artboards") {
                actionHelp.text = "Use the original Group for record 1 on the first new artboard.";
            } else if (flattenStaticCheckbox.value) {
                actionHelp.text = "Replace the template with the first flattened record.";
            } else {
                actionHelp.text = "Use the original Group as record 1.";
            }
        }

        addVariableButton.onClick = addVariable;
        removeVariableButton.onClick = openDeleteVariablesDialog;
        previousMappingButton.onClick = function () {
            syncMappingRowsFromControls();
            if (mappingScrollOffset > 0) { mappingScrollOffset--; }
            refreshMappingPage(true);
        };
        nextMappingButton.onClick = function () {
            syncMappingRowsFromControls();
            if (mappingScrollOffset < Math.max(0, mappingRows.length - mappingRowsPerPage)) { mappingScrollOffset++; }
            refreshMappingPage(true);
        };
        mappingVerticalScrollbar.onChanging = function () {
            var maximumOffset = Math.max(0, mappingRows.length - mappingRowsPerPage);
            var nextOffset = Math.round(Number(mappingVerticalScrollbar.value));
            if (nextOffset < 0) { nextOffset = 0; }
            if (nextOffset > maximumOffset) { nextOffset = maximumOffset; }
            if (nextOffset !== mappingScrollOffset) {
                syncMappingRowsFromControls();
                mappingScrollOffset = nextOffset;
                refreshMappingPage(true);
            }
        };
        mappingVerticalScrollbar.onChange = mappingVerticalScrollbar.onChanging;
        removeVariableButton.enabled = false;
        layoutDropdown.onChange = refreshLayoutControls;
        artboardArrangementDropdown.onChange = refreshLayoutControls;
        actionDropdown.onChange = refreshActionHelp;
        var initialSlotIndex;
        for (initialSlotIndex = 0; initialSlotIndex < mappingRowsPerPage; initialSlotIndex++) {
            createMappingRow(initialSlotIndex);
        }
        createMappingModel("V1", -1);
        refreshMappingPage(true);
        if (initialPreset) {
            try { applyPresetToControls(initialPreset); }
            catch (initialPresetApplyError) {
                alert("The preset was opened, but some settings could not be applied.\n\n" + errorMessage(initialPresetApplyError), SCRIPT_NAME + " - Load Preset");
            }
        }
        refreshMappingPage(true);
        refreshLayoutControls();

        continueButton.onClick = function () {
            try {
                var mappings = readMappings();
                var testConfig = readConfigurationFromControls(
                    layoutDropdown, actionDropdown, columnsInput, rowsInput,
                    horizontalGridSpacingInput, verticalGridSpacingInput,
                    horizontalSpacingInput, verticalSpacingInput,
                    flattenStaticCheckbox, rasterResolutionDropdown, artboardArrangementDropdown
                );
                testConfig.variables = mappings.variables;
                testConfig.frameIndices = mappings.frameIndices;
                validateConfiguration(testConfig, 1);
                win.close(1);
            } catch (setupError) {
                alert(errorMessage(setupError), SCRIPT_NAME + " - Setup Validation");
            }
        };
        cancelButton.onClick = function () { win.close(0); };

        win.readMappings = readMappings;
        win.mappingRows = mappingRows;
        win.layoutDropdown = layoutDropdown;
        win.artboardArrangementDropdown = artboardArrangementDropdown;
        win.actionDropdown = actionDropdown;
        win.columnsInput = columnsInput;
        win.rowsInput = rowsInput;
        win.horizontalGridSpacingInput = horizontalGridSpacingInput;
        win.verticalGridSpacingInput = verticalGridSpacingInput;
        win.horizontalSpacingInput = horizontalSpacingInput;
        win.verticalSpacingInput = verticalSpacingInput;
        win.flattenStaticCheckbox = flattenStaticCheckbox;
        win.rasterResolutionDropdown = rasterResolutionDropdown;
        win.getLoadedPreset = function () { return loadedPreset; };
        return win;
    }

    function buildDataDialog(variableNames, initialRecords, initialBlankBehavior, initialColumnWidths, presetContext) {
        var win = new Window("dialog", SCRIPT_NAME + " " + SCRIPT_VERSION + " - 2/2: Record Data");
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.spacing = 4;
        win.margins = 9;
        win.preferredSize = [820, 435];
        win.maximumSize = [900, 455];

        var intro = win.add("statictext", undefined,
            "Edit records in the grid. Each row is one record. Blank cells are allowed.",
            { multiline: true }
        );
        intro.preferredSize.height = 20;

        var records = [];
        // v1.7 always uses automatic field widths; prior saved manual widths are intentionally ignored.
        var manualColumnWidths = [];
        if (initialRecords && initialRecords.length) {
            var initialRowIndex;
            for (initialRowIndex = 0; initialRowIndex < initialRecords.length; initialRowIndex++) {
                var initialRow = makeBlankRecord(variableNames.length);
                var initialColIndex;
                for (initialColIndex = 0; initialColIndex < variableNames.length; initialColIndex++) {
                    initialRow[initialColIndex] = initialRecords[initialRowIndex][initialColIndex] === undefined
                        ? "" : String(initialRecords[initialRowIndex][initialColIndex]);
                }
                records.push(initialRow);
            }
        }
        var rowPageIndex = 0;
        var variablePageIndex = 0;
        var rowsPerPage = 50;
        var visibleRows = 6;
        var scrollOffset = 0;
        var visibleCellControls = [];
        var rowBindings = [];
        var currentVariableRange = { start: 0, end: 0 };
        var TABLE_CONTENT_WIDTH = 700;
        var TABLE_VIEWPORT_WIDTH = 750;
        var gridBody;
        var gridSpacer;
        var verticalScrollbar;
        var horizontalScrollbar;
        var rowPageLabel;
        var previousRowsButton;
        var nextRowsButton;
        var variablePageLabel;
        var tableHeader;
        var gridViewport;
        var recordCountLabel;

        var tablePanel = win.add("panel", undefined, "Records");
        tablePanel.orientation = "column";
        tablePanel.alignChildren = ["left", "top"];
        tablePanel.margins = 6;
        tablePanel.spacing = 3;
        tablePanel.preferredSize = [790, 250];
        tablePanel.maximumSize = [890, 250];

        tableHeader = tablePanel.add("group");
        tableHeader.orientation = "row";
        tableHeader.alignChildren = ["left", "center"];
        tableHeader.alignment = ["left", "top"];
        tableHeader.spacing = 2;
        tableHeader.preferredSize = [TABLE_CONTENT_WIDTH, 20];
        tableHeader.maximumSize = [TABLE_CONTENT_WIDTH, 20];

        gridViewport = tablePanel.add("group");
        gridViewport.orientation = "row";
        gridViewport.alignChildren = ["left", "top"];
        gridViewport.alignment = ["left", "top"];
        gridViewport.spacing = 3;
        gridViewport.preferredSize = [TABLE_VIEWPORT_WIDTH, 148];
        gridViewport.maximumSize = [TABLE_VIEWPORT_WIDTH, 148];

        gridBody = gridViewport.add("group");
        gridBody.orientation = "column";
        gridBody.alignChildren = ["left", "top"];
        gridBody.alignment = ["left", "top"];
        gridBody.spacing = 1;
        gridBody.preferredSize = [TABLE_CONTENT_WIDTH, 148];
        gridBody.maximumSize = [TABLE_CONTENT_WIDTH, 148];

        gridSpacer = gridViewport.add("group");
        gridSpacer.orientation = "column";
        gridSpacer.preferredSize = [0, 148];
        gridSpacer.maximumSize = [TABLE_CONTENT_WIDTH, 148];

        verticalScrollbar = gridViewport.add("scrollbar", [0, 0, 16, 148], 0, 0, 1);
        verticalScrollbar.preferredSize = [16, 148];
        verticalScrollbar.minimumSize = [16, 148];
        verticalScrollbar.maximumSize = [16, 148];
        verticalScrollbar.alignment = ["left", "top"];
        verticalScrollbar.helpTip = "Scroll through 50 records per page.";

        horizontalScrollbar = tablePanel.add("scrollbar", [0, 0, TABLE_CONTENT_WIDTH, 16], 0, 0, 0);
        horizontalScrollbar.preferredSize = [TABLE_CONTENT_WIDTH, 14];
        horizontalScrollbar.minimumSize = [TABLE_CONTENT_WIDTH, 14];
        horizontalScrollbar.maximumSize = [TABLE_CONTENT_WIDTH, 14];
        horizontalScrollbar.helpTip = "Scroll horizontally through variable columns.";

        var navigation = tablePanel.add("group");
        navigation.orientation = "row";
        navigation.alignChildren = ["center", "center"];
        navigation.alignment = ["left", "top"];
        navigation.spacing = 4;
        previousRowsButton = navigation.add("button", undefined, "< Previous 50");
        rowPageLabel = navigation.add("statictext", undefined, "Page 1/1 | Showing 0-0");
        rowPageLabel.preferredSize.width = 190;
        nextRowsButton = navigation.add("button", undefined, "Next 50 >");
        variablePageLabel = navigation.add("statictext", undefined, "Fields 1-1/1");
        variablePageLabel.preferredSize.width = 115;

        var rowButtons = win.add("group");
        rowButtons.orientation = "row";
        rowButtons.alignChildren = ["left", "center"];
        rowButtons.alignment = ["left", "top"];
        rowButtons.spacing = 4;
        var addRowButton = rowButtons.add("button", undefined, "Add Row +");
        var removeRowButton = rowButtons.add("button", undefined, "Delete Rows...");
        var pasteButton = rowButtons.add("button", undefined, "Paste from Excel...");
        var importFileButton = rowButtons.add("button", undefined, "Import File...");
        recordCountLabel = rowButtons.add("statictext", undefined, "Records: 0");

        // Column widths are automatic: they follow the longest visible cell value.
        var fieldWidthInput = null;
        var applyWidthButton = null;
        var autoWidthButton = null;

        var blankBehaviorRow = win.add("group");
        blankBehaviorRow.orientation = "row";
        blankBehaviorRow.alignChildren = ["center", "center"];
        blankBehaviorRow.alignment = ["left", "top"];
        blankBehaviorRow.spacing = 5;
        blankBehaviorRow.add("statictext", undefined, "Blank cells:");
        var blankBehaviorDropdown = blankBehaviorRow.add("dropdownlist", undefined,
            ["Clear mapped text", "Keep template text"]
        );
        blankBehaviorDropdown.selection = (initialBlankBehavior === "Keep template text") ? 1 : 0;
        blankBehaviorDropdown.preferredSize.width = 165;

        var footer = win.add("statictext", undefined,
            "This Script Made By Ibraheem Sabbah - www.linkedin.com/in/ibraheem-sabbah",
            { multiline: false }
        );
        footer.preferredSize.height = 16;

        var footerNavigation = win.add("group");
        footerNavigation.orientation = "row";
        footerNavigation.alignChildren = ["center", "center"];
        footerNavigation.alignment = ["fill", "top"];
        footerNavigation.spacing = 6;
        var backButton = footerNavigation.add("button", undefined, "< Back to Setup");
        backButton.alignment = ["left", "center"];
        var footerSpacer = footerNavigation.add("statictext", undefined, "");
        footerSpacer.preferredSize = [12, 1];
        footerSpacer.alignment = ["fill", "center"];
        var bottomButtons = footerNavigation.add("group");
        bottomButtons.orientation = "row";
        bottomButtons.alignment = ["right", "center"];
        bottomButtons.alignChildren = ["center", "center"];
        bottomButtons.spacing = 5;
        var previewButton = bottomButtons.add("button", undefined, "Preview #1...");
        var exportPresetButton = bottomButtons.add("button", undefined, "Export Preset...");
        var cancelButton = bottomButtons.add("button", undefined, "Cancel", { name: "cancel" });
        var exportPdfButton = bottomButtons.add("button", undefined, "Export PDF...");
        var finishButton = bottomButtons.add("button", undefined, "Finish", { name: "ok" });
        var requestedPdfMode = "prompt";

        function makeBlankRecord(columnCount) {
            var result = [];
            var j;
            for (j = 0; j < columnCount; j++) { result.push(""); }
            return result;
        }

        function shortLabel(value, maxChars) {
            var s = String(value);
            if (s.length > maxChars) { return s.substring(0, maxChars - 3) + "..."; }
            return s;
        }

        function visibleVariableRange() {
            var start = Math.max(0, Math.min(variablePageIndex, variableNames.length - 1));
            var pageStart = rowPageIndex * rowsPerPage;
            var pageEnd = Math.min(records.length, pageStart + rowsPerPage);
            var budget = TABLE_CONTENT_WIDTH - 38;
            var used = 0;
            var end = start;
            while (end < variableNames.length) {
                var width = calculateCellWidth(end, pageStart, pageEnd);
                if (end === start && width > budget) { width = budget - 2; }
                if (end > start && used + width + 2 > budget) { break; }
                used += width + 2;
                end++;
            }
            if (end === start && variableNames.length > 0) { end = start + 1; }
            return { start: start, end: end };
        }

        function syncVisibleGrid() {
            var i;
            for (i = 0; i < visibleCellControls.length; i++) {
                var control = visibleCellControls[i];
                var recordIndex = control.binding ? control.binding.recordIndex : control.recordIndex;
                if (recordIndex >= 0 && records[recordIndex]) {
                    records[recordIndex][control.variableIndex] = String(control.field.text);
                }
            }
        }

        function clampPageAndScroll() {
            var totalPages = Math.max(1, Math.ceil(records.length / rowsPerPage));
            if (rowPageIndex >= totalPages) { rowPageIndex = totalPages - 1; }
            if (rowPageIndex < 0) { rowPageIndex = 0; }
            var pageStart = rowPageIndex * rowsPerPage;
            var pageCount = Math.max(0, Math.min(rowsPerPage, records.length - pageStart));
            var maxOffset = Math.max(0, pageCount - visibleRows);
            if (scrollOffset > maxOffset) { scrollOffset = maxOffset; }
            if (scrollOffset < 0) { scrollOffset = 0; }
        }

        function calculateCellWidth(variableIndex, pageStart, pageEnd) {
            var maxLen = String(variableNames[variableIndex]).length;
            var i;
            for (i = pageStart; i < pageEnd; i++) {
                var val = records[i] && records[i][variableIndex] !== undefined ? String(records[i][variableIndex]) : "";
                if (val.length > maxLen) { maxLen = val.length; }
            }
            // Estimate a comfortable width from the longest value; cap only at the visible grid width.
            // The text field itself retains horizontal caret scrolling for unusually long values.
            return Math.max(128, Math.min(TABLE_CONTENT_WIDTH - 42, 50 + maxLen * 10.5));
        }

        function updateRowNavigation(firstRecord, lastRecord) {
            var totalRowPages = Math.max(1, Math.ceil(records.length / rowsPerPage));
            var shownFirst = records.length ? firstRecord + 1 : 0;
            rowPageLabel.text = "Page " + (rowPageIndex + 1) + "/" + totalRowPages + " | Showing " + shownFirst + "-" + lastRecord + " of " + records.length;
            previousRowsButton.enabled = rowPageIndex > 0;
            nextRowsButton.enabled = rowPageIndex < totalRowPages - 1;
            recordCountLabel.text = "Records: " + records.length;
        }

        function configureVerticalScrollbar() {
            var pageStart = rowPageIndex * rowsPerPage;
            var pageCount = Math.max(0, Math.min(rowsPerPage, records.length - pageStart));
            var maxOffset = Math.max(0, pageCount - visibleRows);
            verticalScrollbar.minvalue = 0;
            verticalScrollbar.maxvalue = maxOffset;
            verticalScrollbar.stepdelta = 1;
            verticalScrollbar.jumpdelta = 1;
            verticalScrollbar.value = Math.min(scrollOffset, maxOffset);
            verticalScrollbar.enabled = maxOffset > 0;
        }

        function refreshRowsOnly() {
            var pageStart = rowPageIndex * rowsPerPage;
            var pageEnd = Math.min(records.length, pageStart + rowsPerPage);
            var firstRecord = pageStart + scrollOffset;
            var lastRecord = Math.min(pageEnd, firstRecord + visibleRows);
            var i;
            for (i = 0; i < rowBindings.length; i++) {
                var binding = rowBindings[i];
                var recordIndex = firstRecord + i;
                if (recordIndex >= pageEnd) { recordIndex = -1; }
                binding.recordIndex = recordIndex;
                binding.rowPanel.visible = recordIndex >= 0;
                binding.serialLabel.text = recordIndex >= 0 ? String(recordIndex + 1) : "";
                var j;
                for (j = 0; j < binding.fields.length; j++) {
                    var fieldItem = binding.fields[j];
                    var value = (recordIndex >= 0 && records[recordIndex] && records[recordIndex][fieldItem.variableIndex] !== undefined)
                        ? String(records[recordIndex][fieldItem.variableIndex]) : "";
                    fieldItem.field.text = value;
                    fieldItem.field.enabled = recordIndex >= 0;
                    fieldItem.field.helpTip = recordIndex >= 0 ? ("Record " + (recordIndex + 1) + " - " + variableNames[fieldItem.variableIndex]) : "";
                }
            }
            updateRowNavigation(firstRecord, lastRecord);
            configureVerticalScrollbar();
            win.layout.layout(false);
        }

        function resizeVisibleColumns() {
            var range = visibleVariableRange();
            if (range.start !== currentVariableRange.start || range.end !== currentVariableRange.end) {
                renderGrid();
                return;
            }
            var pageStart = rowPageIndex * rowsPerPage;
            var pageEnd = Math.min(records.length, pageStart + rowsPerPage);
            var widths = [];
            var totalWidth = 36;
            var j;
            for (j = range.start; j < range.end; j++) {
                widths[j] = calculateCellWidth(j, pageStart, pageEnd);
                totalWidth += widths[j] + 2;
            }
            for (j = range.start; j < range.end; j++) {
                var headerControl = tableHeader.children[j - range.start + 1];
                if (headerControl) {
                    headerControl.preferredSize = [widths[j], 20];
                    headerControl.maximumSize = [widths[j], 20];
                }
            }
            var i;
            for (i = 0; i < visibleCellControls.length; i++) {
                var item = visibleCellControls[i];
                var w = widths[item.variableIndex];
                if (w) {
                    item.field.preferredSize = [w, 22];
                    item.field.minimumSize = [w, 22];
                    item.field.maximumSize = [w, 22];
                }
            }
            tableHeader.preferredSize = [totalWidth, 20];
            tableHeader.maximumSize = [totalWidth, 20];
            gridBody.preferredSize = [totalWidth, 148];
            gridBody.maximumSize = [totalWidth, 148];
            gridSpacer.preferredSize = [Math.max(0, TABLE_CONTENT_WIDTH - totalWidth - 3), 148];
            gridSpacer.maximumSize = [Math.max(0, TABLE_CONTENT_WIDTH - totalWidth - 3), 148];
            for (i = 0; i < gridBody.children.length; i++) {
                try {
                    gridBody.children[i].preferredSize = [totalWidth, 23];
                    gridBody.children[i].maximumSize = [totalWidth, 23];
                } catch (ignoreResizeRow) {}
            }
            win.layout.layout(false);
        }

        function renderGrid(skipSync) {
            // When deleting rows, skip writing the old controls back into the compacted records array.
            if (!skipSync) { syncVisibleGrid(); }
            clampPageAndScroll();
            while (tableHeader.children.length > 0) { tableHeader.remove(tableHeader.children[0]); }
            while (gridBody.children.length > 0) { gridBody.remove(gridBody.children[0]); }
            visibleCellControls = [];
            rowBindings = [];

            var pageStart = rowPageIndex * rowsPerPage;
            var pageEnd = Math.min(records.length, pageStart + rowsPerPage);
            var totalContentWidth = 36;
            var allWidths = [];
            var j;
            for (j = 0; j < variableNames.length; j++) {
                allWidths[j] = calculateCellWidth(j, pageStart, pageEnd);
                totalContentWidth += allWidths[j] + 2;
            }
            if (totalContentWidth <= TABLE_CONTENT_WIDTH) {
                variablePageIndex = 0;
            } else {
                variablePageIndex = Math.max(0, Math.min(variablePageIndex, variableNames.length - 1));
            }
            var range = visibleVariableRange();
            currentVariableRange = { start: range.start, end: range.end };
            var colWidths = [];
            var tableWidth = 36;
            for (j = range.start; j < range.end; j++) {
                colWidths[j] = allWidths[j];
                tableWidth += colWidths[j] + 2;
            }

            tableHeader.preferredSize = [tableWidth, 20];
            tableHeader.maximumSize = [tableWidth, 20];
            var serialHeader = tableHeader.add("statictext", undefined, "#");
            serialHeader.preferredSize = [32, 20];
            serialHeader.maximumSize = [32, 20];
            serialHeader.justify = "center";
            for (j = range.start; j < range.end; j++) {
                var heading = tableHeader.add("statictext", undefined, shortLabel(variableNames[j], 20));
                heading.preferredSize = [colWidths[j], 20];
                heading.maximumSize = [colWidths[j], 20];
                heading.justify = "center";
                heading.helpTip = variableNames[j];
            }

            gridBody.preferredSize = [tableWidth, 148];
            gridBody.maximumSize = [tableWidth, 148];
            gridSpacer.preferredSize = [Math.max(0, TABLE_CONTENT_WIDTH - tableWidth - 3), 148];
            gridSpacer.maximumSize = [Math.max(0, TABLE_CONTENT_WIDTH - tableWidth - 3), 148];
            var firstRecord = pageStart + scrollOffset;
            var lastRecord = Math.min(pageEnd, firstRecord + visibleRows);
            var i;
            for (i = 0; i < visibleRows; i++) {
                var recordIndex = firstRecord + i;
                if (recordIndex >= pageEnd) { recordIndex = -1; }
                var rowPanel = gridBody.add("panel", undefined, "");
                rowPanel.orientation = "row";
                rowPanel.alignChildren = ["left", "center"];
                rowPanel.alignment = ["left", "top"];
                rowPanel.margins = [2, 1, 2, 1];
                rowPanel.spacing = 2;
                rowPanel.preferredSize = [tableWidth, 23];
                rowPanel.maximumSize = [tableWidth, 23];
                rowPanel.visible = recordIndex >= 0;
                var serialLabel = rowPanel.add("statictext", undefined, recordIndex >= 0 ? String(recordIndex + 1) : "");
                serialLabel.preferredSize = [32, 22];
                serialLabel.maximumSize = [32, 22];
                serialLabel.justify = "center";
                var binding = { recordIndex: recordIndex, rowPanel: rowPanel, serialLabel: serialLabel, fields: [] };
                for (j = range.start; j < range.end; j++) {
                    var initialValue = (recordIndex >= 0 && records[recordIndex] && records[recordIndex][j] !== undefined)
                        ? String(records[recordIndex][j]) : "";
                    var field = rowPanel.add("edittext", undefined, initialValue);
                    field.preferredSize = [colWidths[j], 22];
                    field.minimumSize = [colWidths[j], 22];
                    field.maximumSize = [colWidths[j], 22];
                    field.helpTip = "Record " + (recordIndex + 1) + " - " + variableNames[j];
                    field.enabled = recordIndex >= 0;
                    (function (rowBinding, variableIndex, editField) {
                        editField.onChange = function () {
                            if (rowBinding.recordIndex >= 0 && records[rowBinding.recordIndex]) {
                                records[rowBinding.recordIndex][variableIndex] = String(editField.text);
                            }
                            recordCountLabel.text = "Records: " + records.length;
                            resizeVisibleColumns();
                        };
                    }(binding, j, field));
                    binding.fields.push({ variableIndex: j, field: field });
                    visibleCellControls.push({ binding: binding, variableIndex: j, field: field });
                }
                rowBindings.push(binding);
            }

            updateRowNavigation(firstRecord, lastRecord);
            configureVerticalScrollbar();
            var hasHorizontalOverflow = totalContentWidth > TABLE_CONTENT_WIDTH;
            horizontalScrollbar.minvalue = 0;
            horizontalScrollbar.maxvalue = hasHorizontalOverflow ? Math.max(0, variableNames.length - 1) : 0;
            horizontalScrollbar.stepdelta = 1;
            horizontalScrollbar.jumpdelta = 1;
            horizontalScrollbar.value = variablePageIndex;
            horizontalScrollbar.enabled = horizontalScrollbar.maxvalue > 0;
            variablePageLabel.text = "Fields " + (range.start + 1) + "-" + range.end + "/" + variableNames.length;
            win.layout.layout(true);
        }

        function showLastRecord() {
            var index = Math.max(0, records.length - 1);
            rowPageIndex = Math.floor(index / rowsPerPage);
            var withinPage = index - rowPageIndex * rowsPerPage;
            scrollOffset = Math.max(0, withinPage - visibleRows + 1);
            clampPageAndScroll();
        }

        function addBlankRow() {
            syncVisibleGrid();
            records.push(makeBlankRecord(variableNames.length));
            showLastRecord();
            renderGrid();
        }

        function detectDelimiter(text, preferred) {
            if (preferred) { return preferred; }
            var lines = String(text).replace(/^\uFEFF/, "").split(/\r\n|\n|\r/);
            var line = "";
            var i;
            for (i = 0; i < lines.length; i++) { if (trim(lines[i]) !== "") { line = lines[i]; break; } }
            var counts = { "\t": 0, ",": 0, ";": 0 };
            var quoted = false;
            for (i = 0; i < line.length; i++) {
                var ch = line.charAt(i);
                if (ch === '"') {
                    if (quoted && line.charAt(i + 1) === '"') { i++; }
                    else { quoted = !quoted; }
                } else if (!quoted && counts[ch] !== undefined) { counts[ch]++; }
            }
            if (counts["\t"] >= counts[","] && counts["\t"] >= counts[";"] && counts["\t"] > 0) { return "\t"; }
            if (counts[";"] > counts[","]) { return ";"; }
            return ",";
        }

        function parseDelimitedRows(rawText, preferredDelimiter) {
            var text = String(rawText).replace(/^\uFEFF/, "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
            var delimiter = detectDelimiter(text, preferredDelimiter);
            var rows = [];
            var row = [];
            var cell = "";
            var quoted = false;
            var i;
            for (i = 0; i < text.length; i++) {
                var ch = text.charAt(i);
                if (ch === '"') {
                    if (quoted && text.charAt(i + 1) === '"') { cell += '"'; i++; }
                    else if (quoted) { quoted = false; }
                    else if (cell.length === 0) { quoted = true; }
                    else { cell += ch; }
                } else if (!quoted && ch === delimiter) {
                    row.push(cell); cell = "";
                } else if (!quoted && ch === "\n") {
                    row.push(cell);
                    var blankRow = true;
                    var blankCheckIndex;
                    for (blankCheckIndex = 0; blankCheckIndex < row.length; blankCheckIndex++) {
                        if (trim(String(row[blankCheckIndex])) !== "") { blankRow = false; break; }
                    }
                    if (!blankRow) { rows.push(row); }
                    row = []; cell = "";
                } else {
                    cell += ch;
                }
            }
            if (quoted) { throw new Error("The imported file contains an unclosed quoted field."); }
            row.push(cell);
            var finalBlankRow = true;
            var finalBlankCheckIndex;
            for (finalBlankCheckIndex = 0; finalBlankCheckIndex < row.length; finalBlankCheckIndex++) {
                if (trim(String(row[finalBlankCheckIndex])) !== "") { finalBlankRow = false; break; }
            }
            if (!finalBlankRow) { rows.push(row); }
            return rows;
        }

        function parseImportedRows(rawText, preferredDelimiter) {
            var rawRows = parseDelimitedRows(rawText, preferredDelimiter);
            if (rawRows.length === 0) { throw new Error("No data was found to import. Check that the selected worksheet contains values."); }

            var headerRowIndex = -1;
            var headerMap = [];
            var serialHeaderIndex = -1;
            var usedHeaderIndices = [];
            var scanLimit = Math.min(rawRows.length, 50);
            var i, j, h;

            // Search the first 15 non-empty rows for a header row matching the configured variable names.
            // This supports Excel sheets with a title or blank/metadata rows above the actual table.
            for (i = 0; i < scanLimit; i++) {
                var candidateHeaders = rawRows[i];
                var candidateMap = [];
                var candidateUsed = [];
                var candidateMatches = true;
                for (j = 0; j < variableNames.length; j++) {
                    var wanted = normalizeVariableName(variableNames[j]);
                    var foundAt = -1;
                    for (h = 0; h < candidateHeaders.length; h++) {
                        if (normalizeVariableName(candidateHeaders[h]) === wanted) { foundAt = h; break; }
                    }
                    if (foundAt < 0 || contains(candidateUsed, foundAt)) { candidateMatches = false; break; }
                    candidateMap.push(foundAt);
                    candidateUsed.push(foundAt);
                }
                if (candidateMatches) {
                    headerRowIndex = i;
                    headerMap = candidateMap;
                    usedHeaderIndices = candidateUsed;
                    break;
                }
            }

            var hasHeader = false;
            if (headerRowIndex >= 0) {
                hasHeader = true;
                var foundHeaders = rawRows[headerRowIndex];
                for (h = 0; h < foundHeaders.length; h++) {
                    if (isSerialHeader(foundHeaders[h])) { serialHeaderIndex = h; break; }
                }
                rawRows = rawRows.slice(headerRowIndex + 1);
            } else {
                var firstRow = rawRows[0];
                for (h = 0; h < firstRow.length; h++) {
                    if (isSerialHeader(firstRow[h])) { serialHeaderIndex = h; break; }
                }
                if (serialHeaderIndex >= 0 && firstRow.length >= variableNames.length + 1) {
                    // The header names a serial column, but the variable headings differ from our custom labels.
                    hasHeader = true;
                    rawRows.shift();
                    headerMap = [];
                    var sourceColumn = 0;
                    for (j = 0; j < variableNames.length; j++) {
                        while (sourceColumn === serialHeaderIndex) { sourceColumn++; }
                        headerMap.push(sourceColumn++);
                    }
                } else {
                    // With custom headers (for example Arabic column names), ask whether the first row is a header.
                    if (firstRow.length > 0 && confirm(
                        "No matching variable-name header was found in the first 50 rows.\n\n" +
                        "Does the first row contain column headings that should be skipped?\n\n" +
                        "Yes = skip the first row\nNo = treat rows as data",
                        true,
                        SCRIPT_NAME + " - Import Options"
                    )) {
                        hasHeader = true;
                        rawRows.shift();
                    }
                    headerMap = [];
                    for (j = 0; j < variableNames.length; j++) { headerMap.push(j); }
                }
            }

            var parsed = [];
            for (i = 0; i < rawRows.length; i++) {
                var sourceCells = rawRows[i];
                var rowHasAnyValue = false;
                for (j = 0; j < sourceCells.length; j++) {
                    if (trim(String(sourceCells[j])) !== "") { rowHasAnyValue = true; break; }
                }
                if (!rowHasAnyValue) { continue; }
                var record = makeBlankRecord(variableNames.length);
                for (j = 0; j < variableNames.length; j++) {
                    var sourceIndex = headerMap[j];
                    if (sourceIndex < sourceCells.length) { record[j] = String(sourceCells[sourceIndex]); }
                }

                if (sourceCells.length > variableNames.length) {
                    var mapped = {};
                    for (j = 0; j < headerMap.length; j++) { mapped[String(headerMap[j])] = true; }
                    for (h = 0; h < sourceCells.length; h++) {
                        if (mapped[String(h)]) { continue; }
                        // Allow a recognized serial-number column, but never silently discard other non-empty data.
                        if (serialHeaderIndex === h) { continue; }
                        if (trim(String(sourceCells[h])) !== "") {
                            throw new Error(
                                "The imported data contains more non-empty columns than configured variables.\n\n" +
                                "Remove the extra column, add a matching variable for it, or label the serial-number column as # / Serial / الرقم التسلسلي."
                            );
                        }
                    }
                }
                parsed.push(record);
            }
            if (parsed.length === 0) { throw new Error("No data records were found after skipping headings and empty rows."); }
            return parsed;
        }

        function findExactImportHeader(rawRows) {
            var maxScan = Math.min(rawRows.length, 50);
            var i, j, h;
            for (i = 0; i < maxScan; i++) {
                var headers = rawRows[i];
                var map = [];
                var used = [];
                var matched = true;
                for (j = 0; j < variableNames.length; j++) {
                    var target = normalizeVariableName(variableNames[j]);
                    var found = -1;
                    for (h = 0; h < headers.length; h++) {
                        if (normalizeVariableName(headers[h]) === target) { found = h; break; }
                    }
                    if (found < 0 || contains(used, found)) { matched = false; break; }
                    used.push(found);
                    map.push(found);
                }
                if (matched && map.length === variableNames.length) {
                    return { headerRowIndex: i, sourceColumns: map };
                }
            }
            return null;
        }

        function countSourceColumns(rawRows) {
            var count = 0;
            var i;
            for (i = 0; i < rawRows.length; i++) {
                if (rawRows[i] && rawRows[i].length > count) { count = rawRows[i].length; }
            }
            return count;
        }

        function guessImportColumns(rawRows, headerRowNumber, columnCount) {
            var chosen = [];
            var headers = headerRowNumber > 0 && rawRows[headerRowNumber - 1] ? rawRows[headerRowNumber - 1] : [];
            var used = {};
            var j, h;
            for (j = 0; j < variableNames.length; j++) {
                var found = -1;
                for (h = 0; h < headers.length; h++) {
                    if (!used[String(h)] && normalizeVariableName(headers[h]) === normalizeVariableName(variableNames[j])) {
                        found = h;
                        break;
                    }
                }
                chosen[j] = found;
                if (found >= 0) { used[String(found)] = true; }
            }
            var serialIndex = -1;
            for (h = 0; h < headers.length; h++) {
                if (isSerialHeader(headers[h])) { serialIndex = h; break; }
            }
            var nextColumn = 0;
            for (j = 0; j < variableNames.length; j++) {
                if (chosen[j] >= 0) { continue; }
                while (nextColumn < columnCount && (used[String(nextColumn)] || nextColumn === serialIndex)) { nextColumn++; }
                chosen[j] = nextColumn < columnCount ? nextColumn : -1;
                if (chosen[j] >= 0) { used[String(chosen[j])] = true; nextColumn++; }
            }
            return chosen;
        }

        function mapSourceRows(rawRows, headerRowNumber, sourceColumns) {
            var startIndex = headerRowNumber > 0 ? headerRowNumber : 0;
            var mappedRows = [];
            var i, j;
            for (i = startIndex; i < rawRows.length; i++) {
                var source = rawRows[i] || [];
                var result = makeBlankRecord(variableNames.length);
                var hasMappedValue = false;
                for (j = 0; j < variableNames.length; j++) {
                    var sourceIndex = sourceColumns[j];
                    if (sourceIndex !== undefined && sourceIndex >= 0 && sourceIndex < source.length) {
                        result[j] = String(source[sourceIndex] === undefined ? "" : source[sourceIndex]);
                        if (trim(result[j]) !== "") { hasMappedValue = true; }
                    }
                }
                if (hasMappedValue) { mappedRows.push(result); }
            }
            if (!mappedRows.length) {
                throw new Error("No non-empty records were found in the selected data range. Check the header row and source-column mappings.");
            }
            return mappedRows;
        }

        function parseImportWithColumnMapping(rawText, preferredDelimiter) {
            var rawRows = parseDelimitedRows(rawText, preferredDelimiter);
            if (!rawRows.length) { throw new Error("The selected source contains no readable rows."); }
            var exact = findExactImportHeader(rawRows);
            if (exact) {
                // Exact variable-name headers are mapped automatically; extra columns such as serial numbers are ignored.
                return mapSourceRows(rawRows, exact.headerRowIndex + 1, exact.sourceColumns);
            }
            // Never silently discard the first data row just because names contain text.
            // Ask explicitly when custom headers do not exactly match the target variable names.
            var hasHeaderRow = confirm(
                "The file's column headings do not exactly match the variable names.\n\n" +
                "Does the first row contain column headings that should be skipped?\n\n" +
                "Yes = use row 1 as headings\nNo = keep row 1 as data",
                true,
                SCRIPT_NAME + " - Import Options"
            );
            return openImportColumnMappingDialog(rawRows, hasHeaderRow ? 1 : 0);
        }

        function openImportColumnMappingDialog(rawRows, headerRowDefault) {
            var columnCount = countSourceColumns(rawRows);
            if (columnCount < 1) { throw new Error("No source columns were found in the selected data."); }
            if (headerRowDefault !== 0 && headerRowDefault !== 1) { headerRowDefault = 0; }

            var importWin = new Window("dialog", SCRIPT_NAME + " - Import Column Mapping");
            importWin.orientation = "column";
            importWin.alignChildren = ["fill", "top"];
            importWin.spacing = 5;
            importWin.margins = 10;
            importWin.preferredSize = [610, 355];
            importWin.add("statictext", undefined,
                "Match each variable to a source column. Unused source columns are ignored.",
                { multiline: true }
            ).preferredSize.height = 22;

            var headerControlRow = importWin.add("group");
            headerControlRow.orientation = "row";
            headerControlRow.alignChildren = ["center", "center"];
            headerControlRow.add("statictext", undefined, "Header row (0 = no header):");
            var headerRowInput = headerControlRow.add("edittext", undefined, String(headerRowDefault));
            headerRowInput.characters = 4;
            var updateHeadersButton = headerControlRow.add("button", undefined, "Refresh Columns");
            var sourceInfo = importWin.add("statictext", undefined, "Source columns: " + columnCount);
            sourceInfo.preferredSize.height = 16;

            var mappingHeader = importWin.add("group");
            mappingHeader.orientation = "row";
            mappingHeader.alignChildren = ["center", "center"];
            mappingHeader.add("statictext", undefined, "Target variable").preferredSize.width = 190;
            mappingHeader.add("statictext", undefined, "Source column").preferredSize.width = 350;

            var mappingContainer = importWin.add("group");
            mappingContainer.orientation = "column";
            mappingContainer.alignChildren = ["fill", "top"];
            mappingContainer.spacing = 2;
            var nav = importWin.add("group");
            nav.orientation = "row";
            nav.alignChildren = ["center", "center"];
            var previous = nav.add("button", undefined, "< Previous Fields");
            var pageLabel = nav.add("statictext", undefined, "Fields 1-1");
            pageLabel.preferredSize.width = 140;
            var next = nav.add("button", undefined, "Next Fields >");
            var buttonRow = importWin.add("group");
            buttonRow.orientation = "row";
            buttonRow.alignment = ["right", "top"];
            var cancel = buttonRow.add("button", undefined, "Cancel", { name: "cancel" });
            var ok = buttonRow.add("button", undefined, "Import Records", { name: "ok" });

            var mapControls = [];
            var mapPageIndex = 0;
            var mapRowsPerPage = 5;
            var selectedColumns = guessImportColumns(rawRows, headerRowDefault, columnCount);
            var resultRows = null;

            function validHeaderRowNumber() {
                var s = trim(String(headerRowInput.text));
                if (!/^\d+$/.test(s)) { throw new Error("Header row must be 0 or a positive whole number."); }
                var n = parseInt(s, 10);
                if (n > rawRows.length) { throw new Error("Header row cannot exceed the source's " + rawRows.length + " non-empty row(s)."); }
                return n;
            }

            function sourceColumnLabels(headerRowNumber) {
                var labels = [];
                var headers = headerRowNumber > 0 && rawRows[headerRowNumber - 1] ? rawRows[headerRowNumber - 1] : [];
                var i;
                for (i = 0; i < columnCount; i++) {
                    var label = headerRowNumber > 0 && headers[i] !== undefined ? trim(String(headers[i])) : "";
                    if (!label) { label = "Column " + (i + 1); }
                    label = String.fromCharCode(65 + (i % 26)) + " | " + label;
                    labels.push(label);
                }
                return labels;
            }

            function rebuildColumnChoices(resetDefaults) {
                var headerRowNumber = validHeaderRowNumber();
                var labels = sourceColumnLabels(headerRowNumber);
                if (resetDefaults) { selectedColumns = guessImportColumns(rawRows, headerRowNumber, columnCount); }
                var i;
                for (i = 0; i < mapControls.length; i++) {
                    var dd = mapControls[i].dropdown;
                    var prior = selectedColumns[i];
                    try { dd.removeAll(); } catch (removeAllError) {
                        while (dd.items.length > 0) { dd.remove(dd.items[0]); }
                    }
                    dd.add("item", "(Leave blank)");
                    var c;
                    for (c = 0; c < labels.length; c++) { dd.add("item", labels[c]); }
                    if (prior < -1 || prior >= columnCount) { prior = -1; }
                    dd.selection = prior < 0 ? 0 : prior + 1;
                }
                var first = mapPageIndex * mapRowsPerPage;
                var last = Math.min(mapControls.length, first + mapRowsPerPage);
                for (i = 0; i < mapControls.length; i++) { mapControls[i].group.visible = (i >= first && i < last); }
                pageLabel.text = "Fields " + (mapControls.length ? first + 1 : 0) + "-" + last + " of " + mapControls.length;
                previous.enabled = mapPageIndex > 0;
                next.enabled = mapPageIndex < Math.ceil(mapControls.length / mapRowsPerPage) - 1;
                sourceInfo.text = "Source columns: " + columnCount + " | Data rows start after header row " + headerRowNumber;
                importWin.layout.layout(true);
            }

            var i;
            for (i = 0; i < variableNames.length; i++) {
                var row = mappingContainer.add("group");
                row.orientation = "row";
                row.alignChildren = ["center", "center"];
                var targetLabel = row.add("statictext", undefined, variableNames[i]);
                targetLabel.preferredSize.width = 190;
                var dd = row.add("dropdownlist", undefined, []);
                dd.preferredSize.width = 350;
                mapControls.push({ group: row, dropdown: dd });
            }
            function syncChosenColumns() {
                var k;
                for (k = 0; k < mapControls.length; k++) {
                    selectedColumns[k] = mapControls[k].dropdown.selection ? mapControls[k].dropdown.selection.index - 1 : -1;
                }
            }
            updateHeadersButton.onClick = function () {
                try { syncChosenColumns(); rebuildColumnChoices(true); }
                catch (headerError) { alert(errorMessage(headerError), SCRIPT_NAME + " - Import Mapping"); }
            };
            previous.onClick = function () {
                syncChosenColumns();
                if (mapPageIndex > 0) { mapPageIndex--; }
                rebuildColumnChoices(false);
            };
            next.onClick = function () {
                syncChosenColumns();
                if ((mapPageIndex + 1) * mapRowsPerPage < mapControls.length) { mapPageIndex++; }
                rebuildColumnChoices(false);
            };
            cancel.onClick = function () { importWin.close(0); };
            ok.onClick = function () {
                try {
                    syncChosenColumns();
                    var headerRowNumber = validHeaderRowNumber();
                    resultRows = mapSourceRows(rawRows, headerRowNumber, selectedColumns);
                    importWin.close(1);
                } catch (mapError) { alert(errorMessage(mapError), SCRIPT_NAME + " - Import Mapping"); }
            };
            rebuildColumnChoices(false);
            if (importWin.show() !== 1) { return null; }
            return resultRows;
        }

        function appendImportedRows(imported, sourceLabel) {
            syncVisibleGrid();
            var firstNewRow = records.length + 1;
            var i;
            for (i = 0; i < imported.length; i++) { records.push(imported[i]); }
            showLastRecord();
            renderGrid();
            alert("Appended " + imported.length + " record(s), starting at row " + firstNewRow + ".\nExisting records were preserved.", sourceLabel);
        }

        function openPasteDialog() {
            var pasteWin = new Window("dialog", "Paste Data from Excel");
            pasteWin.orientation = "column";
            pasteWin.alignChildren = ["fill", "top"];
            pasteWin.spacing = 5;
            pasteWin.margins = 9;
            pasteWin.preferredSize = [660, 300];
            pasteWin.maximumSize = [760, 380];
            pasteWin.add("statictext", undefined,
                "Copy cells from Excel and paste them here. A matching header row is optional.",
                { multiline: true }
            ).preferredSize.height = 22;
            var pasteArea = pasteWin.add("edittext", undefined, "", { multiline: true, scrolling: true, wantReturn: true });
            pasteArea.preferredSize = [635, 205];
            pasteArea.alignment = ["fill", "fill"];
            var pasteButtons = pasteWin.add("group");
            pasteButtons.orientation = "row";
            pasteButtons.alignment = ["right", "top"];
            var pasteCancel = pasteButtons.add("button", undefined, "Cancel", { name: "cancel" });
            var pasteOK = pasteButtons.add("button", undefined, "Append Records", { name: "ok" });
            pasteCancel.onClick = function () { pasteWin.close(0); };
            pasteOK.onClick = function () {
                try {
                    var imported = parseImportWithColumnMapping(pasteArea.text, "\t");
                    if (!imported) { return; }
                    appendImportedRows(imported, SCRIPT_NAME + " - Excel Paste");
                    pasteWin.close(1);
                } catch (pasteError) { alert(errorMessage(pasteError), "Excel Paste"); }
            };
            pasteWin.show();
        }

        function vbStringLiteral(value) {
            return '"' + String(value).replace(/"/g, '""') + '"';
        }

        function getExcelWorksheetNames(file) {
            if (String($.os).toLowerCase().indexOf("windows") < 0) {
                throw new Error("Direct workbook import requires Windows with Microsoft Excel installed. On other systems, export the workbook as CSV UTF-8 or TSV.");
            }
            var token = String(new Date().getTime()) + "_sheets_" + String(Math.floor(Math.random() * 1000000));
            var namesFile = File(Folder.temp.fsName + "/VDS_" + token + ".sheets.tsv");
            var statusFile = File(Folder.temp.fsName + "/VDS_" + token + ".status");
            var logFile = File(Folder.temp.fsName + "/VDS_" + token + ".console.log");
            var helperFile = File(Folder.temp.fsName + "/VDS_" + token + ".vbs");
            var vb = [];
            vb.push("Option Explicit");
            vb.push("Dim inputPath, namesPath, statusPath, xl, wb, stream, i, errorText, lineValue");
            vb.push("inputPath = " + vbStringLiteral(file.fsName));
            vb.push("namesPath = " + vbStringLiteral(namesFile.fsName));
            vb.push("statusPath = " + vbStringLiteral(statusFile.fsName));
            vb.push("Sub WriteUtf8(ByVal filePath, ByVal content)");
            vb.push("  Dim st");
            vb.push("  Set st = CreateObject(\"ADODB.Stream\")");
            vb.push("  st.Type = 2: st.Charset = \"utf-8\": st.Open");
            vb.push("  st.WriteText content");
            vb.push("  st.SaveToFile filePath, 2");
            vb.push("  st.Close");
            vb.push("End Sub");
            vb.push("Sub Fail(ByVal message)");
            vb.push("  On Error Resume Next");
            vb.push("  If Not wb Is Nothing Then wb.Close False");
            vb.push("  If Not xl Is Nothing Then xl.Quit");
            vb.push("  WriteUtf8 statusPath, \"ERROR|\" & message");
            vb.push("  WScript.Quit 2");
            vb.push("End Sub");
            vb.push("On Error Resume Next");
            vb.push("Set xl = CreateObject(\"Excel.Application\")");
            vb.push("If Err.Number <> 0 Then errorText = Err.Description: Err.Clear: Fail \"Could not start Microsoft Excel: \" & errorText");
            vb.push("xl.Visible = False: xl.DisplayAlerts = False: xl.EnableEvents = False: xl.AskToUpdateLinks = False");
            vb.push("xl.AutomationSecurity = 3");
            vb.push("Err.Clear");
            vb.push("Set wb = xl.Workbooks.Open(inputPath, 0, True)");
            vb.push("If Err.Number <> 0 Then errorText = Err.Description: Err.Clear: Fail \"Could not open the workbook read-only: \" & errorText");
            vb.push("Set stream = CreateObject(\"ADODB.Stream\")");
            vb.push("If Err.Number <> 0 Then errorText = Err.Description: Err.Clear: Fail \"Windows text-stream support is unavailable: \" & errorText");
            vb.push("stream.Type = 2: stream.Charset = \"utf-8\": stream.Open");
            vb.push("For i = 1 To wb.Worksheets.Count");
            vb.push("  lineValue = CStr(i) & vbTab & CStr(wb.Worksheets(i).Name)");
            vb.push("  stream.WriteText lineValue & vbCrLf");
            vb.push("Next");
            vb.push("stream.SaveToFile namesPath, 2");
            vb.push("If Err.Number <> 0 Then errorText = Err.Description: Err.Clear: stream.Close: Fail \"Could not list worksheet names: \" & errorText");
            vb.push("stream.Close: wb.Close False: xl.Quit");
            vb.push("WriteUtf8 statusPath, \"OK\"");
            vb.push("WScript.Quit 0");

            helperFile.encoding = "UTF-16";
            if (!helperFile.open("w")) { throw new Error("Could not create the Excel worksheet-list helper."); }
            helperFile.write(vb.join("\r\n"));
            helperFile.close();
            try {
                runWindowsVbsHelper(helperFile, statusFile, logFile, "listing worksheets");
            } catch (sheetLaunchError) {
                try { if (helperFile.exists) { helperFile.remove(); } } catch (ignoreSheetHelperErrorCleanup) {}
                try { if (statusFile.exists) { statusFile.remove(); } } catch (ignoreSheetStatusErrorCleanup) {}
                try { if (logFile.exists) { logFile.remove(); } } catch (ignoreSheetLogErrorCleanup) {}
                try { if (namesFile.exists) { namesFile.remove(); } } catch (ignoreSheetNamesErrorCleanup) {}
                throw sheetLaunchError;
            }
            if (!namesFile.exists) { throw new Error("Excel did not return any worksheet names."); }
            namesFile.encoding = "UTF-8";
            if (!namesFile.open("r")) { throw new Error("Could not read the worksheet-name list."); }
            var rawNames = namesFile.read();
            namesFile.close();
            try { helperFile.remove(); } catch (ignoreSheetHelper) {}
            try { statusFile.remove(); } catch (ignoreSheetStatus) {}
            try { logFile.remove(); } catch (ignoreSheetLog) {}
            try { namesFile.remove(); } catch (ignoreSheetNames) {}
            rawNames = String(rawNames).replace(/^\uFEFF/, "").replace(/\r/g, "");
            var lines = rawNames.split("\n");
            var result = [];
            var lineIndex;
            for (lineIndex = 0; lineIndex < lines.length; lineIndex++) {
                if (trim(lines[lineIndex]) === "") { continue; }
                var splitAt = lines[lineIndex].indexOf("\t");
                if (splitAt < 0) { continue; }
                var sheetIndex = parseInt(lines[lineIndex].substring(0, splitAt), 10);
                var sheetName = lines[lineIndex].substring(splitAt + 1);
                if (sheetIndex > 0) { result.push({ index: sheetIndex, name: sheetName }); }
            }
            if (!result.length) { throw new Error("No worksheets were found in this Excel workbook."); }
            return result;
        }

        function chooseExcelWorksheet(sheetInfos) {
            if (!sheetInfos || !sheetInfos.length) { throw new Error("No worksheet is available to import."); }
            if (sheetInfos.length === 1) { return sheetInfos[0].index; }
            var sheetWin = new Window("dialog", SCRIPT_NAME + " - Choose Worksheet");
            sheetWin.orientation = "column";
            sheetWin.alignChildren = ["fill", "top"];
            sheetWin.spacing = 6;
            sheetWin.margins = 10;
            sheetWin.preferredSize = [360, 125];
            sheetWin.add("statictext", undefined, "Choose the worksheet containing the records:");
            var sheetDropdown = sheetWin.add("dropdownlist", undefined, []);
            var i;
            for (i = 0; i < sheetInfos.length; i++) {
                sheetDropdown.add("item", sheetInfos[i].index + " | " + sheetInfos[i].name);
            }
            sheetDropdown.selection = 0;
            var buttons = sheetWin.add("group");
            buttons.orientation = "row";
            buttons.alignment = ["right", "top"];
            buttons.add("button", undefined, "Cancel", { name: "cancel" }).onClick = function () { sheetWin.close(0); };
            buttons.add("button", undefined, "Import Sheet", { name: "ok" }).onClick = function () { sheetWin.close(1); };
            if (sheetWin.show() !== 1 || !sheetDropdown.selection) { return 0; }
            return sheetInfos[sheetDropdown.selection.index].index;
        }

        function importNativeExcelWorkbook(file, selectedWorksheetIndex) {
            if (String($.os).toLowerCase().indexOf("windows") < 0) {
                throw new Error("Direct Excel workbook import is available only on Windows with Microsoft Excel installed. Export the workbook as CSV UTF-8 or Tab Delimited Text, or use Paste from Excel.");
            }

            // Let Excel export the selected worksheet through its native Unicode Text format.
            // This avoids fragile per-cell COM conversions (especially formula errors and Arabic/localized values).
            var token = String(new Date().getTime()) + "_" + String(Math.floor(Math.random() * 1000000));
            var tempFolder = Folder.temp;
            var outputFile = File(tempFolder.fsName + "/VDS_" + token + ".txt");
            var statusFile = File(tempFolder.fsName + "/VDS_" + token + ".status");
            var logFile = File(tempFolder.fsName + "/VDS_" + token + ".console.log");
            var helperFile = File(tempFolder.fsName + "/VDS_" + token + ".vbs");
            var vb = [];
            vb.push("Option Explicit");
            vb.push("Dim inputPath, outputPath, statusPath, xl, wb, ws, exportWb, selectedWorksheetIndex, errorText");
            vb.push("inputPath = " + vbStringLiteral(file.fsName));
            vb.push("selectedWorksheetIndex = " + String(parseInt(selectedWorksheetIndex, 10) || 0));
            vb.push("outputPath = " + vbStringLiteral(outputFile.fsName));
            vb.push("statusPath = " + vbStringLiteral(statusFile.fsName));
            vb.push("Sub WriteUtf8(ByVal filePath, ByVal content)");
            vb.push("  Dim st");
            vb.push("  Set st = CreateObject(\"ADODB.Stream\")");
            vb.push("  st.Type = 2: st.Charset = \"utf-8\": st.Open");
            vb.push("  st.WriteText content");
            vb.push("  st.SaveToFile filePath, 2");
            vb.push("  st.Close");
            vb.push("End Sub");
            vb.push("Sub Fail(ByVal message)");
            vb.push("  On Error Resume Next");
            vb.push("  If Not exportWb Is Nothing Then exportWb.Close False");
            vb.push("  If Not wb Is Nothing Then wb.Close False");
            vb.push("  If Not xl Is Nothing Then xl.Quit");
            vb.push("  WriteUtf8 statusPath, \"ERROR|\" & message");
            vb.push("  WScript.Quit 2");
            vb.push("End Sub");
            vb.push("On Error Resume Next");
            vb.push("Set xl = CreateObject(\"Excel.Application\")");
            vb.push("If Err.Number <> 0 Then errorText = Err.Description: Err.Clear: Fail \"Could not start Microsoft Excel: \" & errorText");
            vb.push("xl.Visible = False");
            vb.push("xl.DisplayAlerts = False");
            vb.push("xl.EnableEvents = False");
            vb.push("xl.AskToUpdateLinks = False");
            vb.push("Err.Clear");
            vb.push("On Error Resume Next");
            vb.push("xl.AutomationSecurity = 3");
            vb.push("Err.Clear");
            vb.push("Set wb = xl.Workbooks.Open(inputPath, 0, True)");
            vb.push("If Err.Number <> 0 Then errorText = Err.Description: Err.Clear: Fail \"Could not open the workbook read-only: \" & errorText");
            vb.push("If selectedWorksheetIndex < 1 Or selectedWorksheetIndex > wb.Worksheets.Count Then Fail \"Selected worksheet index is out of range.\"");
            vb.push("Set ws = wb.Worksheets(selectedWorksheetIndex)");
            vb.push("If Err.Number <> 0 Then errorText = Err.Description: Err.Clear: Fail \"Could not select the worksheet: \" & errorText");
            vb.push("If xl.WorksheetFunction.CountA(ws.UsedRange) = 0 Then Fail \"The selected worksheet is empty. Choose a worksheet containing records.\"");
            vb.push("Err.Clear");
            vb.push("ws.Copy");
            vb.push("If Err.Number <> 0 Then errorText = Err.Description: Err.Clear: Fail \"Could not copy the selected worksheet for export: \" & errorText");
            vb.push("Set exportWb = xl.ActiveWorkbook");
            vb.push("If exportWb Is Nothing Then Fail \"Excel did not create a temporary workbook for export.\"");
            // Excel constant xlUnicodeText = 42. The selected worksheet is saved as UTF-16 tab-delimited text.
            vb.push("Err.Clear");
            vb.push("exportWb.SaveAs outputPath, 42");
            vb.push("If Err.Number <> 0 Then errorText = Err.Description: Err.Clear: Fail \"Could not save the worksheet as Unicode text: \" & errorText");
            vb.push("exportWb.Close False");
            vb.push("Set exportWb = Nothing");
            vb.push("wb.Close False");
            vb.push("xl.Quit");
            vb.push("WriteUtf8 statusPath, \"OK\"");
            vb.push("WScript.Quit 0");

            helperFile.encoding = "UTF-16";
            if (!helperFile.open("w")) { throw new Error("Could not create the temporary Excel import helper."); }
            helperFile.write(vb.join("\r\n"));
            helperFile.close();

            try {
                runWindowsVbsHelper(helperFile, statusFile, logFile, "importing worksheet");
            } catch (importLaunchError) {
                try { if (helperFile.exists) { helperFile.remove(); } } catch (ignoreImportHelperErrorCleanup) {}
                try { if (statusFile.exists) { statusFile.remove(); } } catch (ignoreImportStatusErrorCleanup) {}
                try { if (logFile.exists) { logFile.remove(); } } catch (ignoreImportLogErrorCleanup) {}
                try { if (outputFile.exists) { outputFile.remove(); } } catch (ignoreImportOutputErrorCleanup) {}
                throw importLaunchError;
            }

            if (!outputFile.exists) {
                throw new Error("Excel did not produce the Unicode worksheet export. Close any Excel dialogs and try again.");
            }

            // Native Unicode Text export is UTF-16 and tab-delimited.
            outputFile.encoding = "UTF-16";
            if (!outputFile.open("r")) { throw new Error("Could not read the temporary Unicode worksheet export."); }
            var contents = outputFile.read();
            outputFile.close();
            try { helperFile.remove(); } catch (ignoreHelperCleanup) {}
            try { statusFile.remove(); } catch (ignoreStatusCleanup) {}
            try { logFile.remove(); } catch (ignoreImportLogCleanup) {}
            try { outputFile.remove(); } catch (ignoreOutputCleanup) {}
            return String(contents).replace(/^\uFEFF/, "");
        }

        function openImportFileDialog() {
            // Do not use a ScriptUI extension callback here: it filters out valid workbook
            // choices in some Illustrator/Windows combinations. Let the OS show all files,
            // then validate the selected extension below.
            var file = File.openDialog(
                "Select an Excel workbook or data file (.xlsx, .xls, .xlsm, .csv, .tsv, .txt)"
            );
            if (!file) { return; }
            var extensionMatch = String(file.name).match(/\.([^.]+)$/);
            var extension = extensionMatch ? extensionMatch[1].toLowerCase() : "";
            if (!/^(csv|tsv|txt|xlsx|xls|xlsm)$/.test(extension)) {
                alert("Unsupported file type. Choose .xlsx, .xls, .xlsm, .csv, .tsv, or .txt.", SCRIPT_NAME + " - File Import");
                return;
            }
            try {
                var contents;
                var delimiter;
                if (extension === "xlsx" || extension === "xls" || extension === "xlsm") {
                    var worksheetInfos = getExcelWorksheetNames(file);
                    var selectedWorksheet = chooseExcelWorksheet(worksheetInfos);
                    if (!selectedWorksheet) { return; }
                    contents = importNativeExcelWorkbook(file, selectedWorksheet);
                    delimiter = "\t";
                } else {
                    file.encoding = "UTF-8";
                    if (!file.open("r")) { throw new Error("Could not open the selected file."); }
                    contents = file.read();
                    file.close();
                    delimiter = extension === "tsv" || extension === "txt" ? "\t" : null;
                }
                var imported = parseImportWithColumnMapping(contents, delimiter);
                if (!imported) { return; }
                appendImportedRows(imported, SCRIPT_NAME + " - File Import");
            } catch (fileError) {
                try { if (file.opened) { file.close(); } } catch (ignoreFileClose) {}
                alert(errorMessage(fileError), SCRIPT_NAME + " - File Import");
            }
        }

        function parseRowDeletionSpec(spec, totalRecords) {
            return parseIndexSpecification(spec, totalRecords, "Row");
        }

        function openDeleteRowsDialog() {
            syncVisibleGrid();
            if (records.length < 1) { return; }
            var deleteWin = new Window("dialog", SCRIPT_NAME + " - Delete Rows");
            deleteWin.orientation = "column";
            deleteWin.alignChildren = ["fill", "top"];
            deleteWin.spacing = 8;
            deleteWin.margins = 12;
            deleteWin.preferredSize = [390, 125];
            deleteWin.add("statictext", undefined, "Rows to delete (current serial numbers):");
            var rowSpecInput = deleteWin.add("edittext", undefined, "");
            rowSpecInput.characters = 36;
            rowSpecInput.active = true;
            var deleteButtons = deleteWin.add("group");
            deleteButtons.orientation = "row";
            deleteButtons.alignment = ["right", "top"];
            var deleteCancelButton = deleteButtons.add("button", undefined, "Cancel", { name: "cancel" });
            var deleteOKButton = deleteButtons.add("button", undefined, "Delete", { name: "ok" });
            var parsedRowNumbers = null;
            deleteCancelButton.onClick = function () { deleteWin.close(0); };
            deleteOKButton.onClick = function () {
                try {
                    parsedRowNumbers = parseRowDeletionSpec(rowSpecInput.text, records.length);
                    if (!parsedRowNumbers.length) { throw new Error("Enter at least one row number."); }
                    deleteWin.close(1);
                } catch (deleteParseError) {
                    alert(errorMessage(deleteParseError), SCRIPT_NAME + " - Delete Rows");
                }
            };
            if (deleteWin.show() !== 1 || !parsedRowNumbers) { return; }
            if (!confirm("Delete " + parsedRowNumbers.length + " selected row(s)?", true)) { return; }

            // Filter by the displayed one-based serial numbers through one tested helper.
            // This avoids using moving row-binding indices after the array is compacted.
            records = filterItemsByIndexSpecification(records, rowSpecInput.text, "Row");
            // Invalidate bindings before rebuilding. Old controls still reference the pre-delete row indexes.
            visibleCellControls = [];
            rowBindings = [];
            clampPageAndScroll();
            renderGrid(true);
        }

        function serializeData() {
            syncVisibleGrid();
            var lines = [variableNames.join("\t")];
            var i;
            var j;
            for (i = 0; i < records.length; i++) {
                var cells = [];
                for (j = 0; j < variableNames.length; j++) {
                    cells.push(String(records[i][j] === undefined ? "" : records[i][j]));
                }
                lines.push(cells.join("\t"));
            }
            return lines.join("\n");
        }

        previousRowsButton.onClick = function () {
            syncVisibleGrid();
            if (rowPageIndex > 0) { rowPageIndex--; }
            scrollOffset = 0;
            renderGrid();
        };
        nextRowsButton.onClick = function () {
            syncVisibleGrid();
            if ((rowPageIndex + 1) * rowsPerPage < records.length) { rowPageIndex++; }
            scrollOffset = 0;
            renderGrid();
        };
        horizontalScrollbar.onChanging = function () {
            var nextVariableIndex = Math.round(Number(horizontalScrollbar.value));
            nextVariableIndex = Math.max(0, Math.min(variableNames.length - 1, nextVariableIndex));
            if (nextVariableIndex !== variablePageIndex) {
                syncVisibleGrid();
                variablePageIndex = nextVariableIndex;
                // The current values were synced just once; skip renderGrid's second sync to keep dragging responsive.
                renderGrid(true);
            }
        };
        horizontalScrollbar.onChange = horizontalScrollbar.onChanging;
        verticalScrollbar.onChanging = function () {
            var nextOffset = Math.round(Number(verticalScrollbar.value));
            if (nextOffset !== scrollOffset) {
                syncVisibleGrid();
                scrollOffset = nextOffset;
                refreshRowsOnly();
            }
        };
        verticalScrollbar.onChange = verticalScrollbar.onChanging;
        addRowButton.onClick = addBlankRow;
        removeRowButton.onClick = openDeleteRowsDialog;
        pasteButton.onClick = openPasteDialog;
        importFileButton.onClick = openImportFileDialog;
        exportPresetButton.onClick = function () {
            syncVisibleGrid();
            try {
                var exportFile = File.saveDialog("Export all settings and records as a preset", "Variable Data Studio Preset:*.vds");
                if (!exportFile) { return; }
                if (!/\.vds$/i.test(exportFile.name)) { exportFile = new File(exportFile.fsName + ".vds"); }
                var result = savePresetPackage(exportFile, exportFile.name.replace(/\.(?:vds|cbg)$/i, ""), presetContext, variableNames, records, getBlankBehaviorValue(), manualColumnWidths, true);
                rememberLastPresetPath(exportFile);
                var message = "Preset exported successfully:\n" + exportFile.fsName;
                if (result.templateBackupPath) { message += "\n\nTemplate copy:\n" + result.templateBackupPath; }
                if (result.warning) { message += "\n\nNote: " + result.warning; }
                alert(message, SCRIPT_NAME + " - Export Preset");
            } catch (exportError) {
                alert(errorMessage(exportError), SCRIPT_NAME + " - Export Preset");
            }
        };
        function saveCurrentPreset() {
            var defaultPresetName = presetContext.groupName ? String(presetContext.groupName) : "Data Preset";
            var presetFile = null;
            var presetName = "";
            var useAutoLibrary = confirm(
                "Save automatically to the Variable Data Studio Library?\n\nYes = Auto Library\nNo = Choose a file location",
                true,
                SCRIPT_NAME + " - Save Preset"
            );
            if (useAutoLibrary) {
                var requestedName = prompt("Name this preset:", defaultPresetName, SCRIPT_NAME + " - Save Preset");
                if (requestedName === null || trim(String(requestedName)) === "") { return false; }
                presetName = trim(String(requestedName));
                var libraryFolder = getPresetLibraryFolder();
                var safeName = sanitizePresetFileName(presetName);
                presetFile = new File(libraryFolder.fsName + "/" + safeName + ".vds");
            } else {
                presetFile = File.saveDialog("Choose where to save the preset (.vds)", "Variable Data Studio Preset:*.vds");
                if (!presetFile) { return false; }
                if (!/\.vds$/i.test(presetFile.name)) { presetFile = new File(presetFile.fsName + ".vds"); }
                presetName = presetFile.name.replace(/\.(?:vds|cbg)$/i, "");
            }
            if (presetFile.exists && !confirm("A preset with this name already exists. Overwrite it?", true)) { return false; }
            var saveResult = savePresetPackage(presetFile, presetName, presetContext, variableNames, records, getBlankBehaviorValue(), manualColumnWidths, true);
            rememberLastPresetPath(presetFile);
            var saveMessage = "Preset saved successfully:\n" + presetFile.fsName;
            if (saveResult.templateBackupPath) { saveMessage += "\n\nOriginal template copy:\n" + saveResult.templateBackupPath; }
            if (saveResult.warning) { saveMessage += "\n\nNote: " + saveResult.warning; }
            alert(saveMessage, SCRIPT_NAME + " - Save Preset");
            return true;
        }

        previewButton.onClick = function () {
            try {
                syncVisibleGrid();
                if (!records.length) {
                    alert("Add or import at least one record before previewing.", SCRIPT_NAME + " - Preview");
                    return;
                }
                var previewRow = makeDataRowsFromGrid([records[0]], variableNames).rows[0];
                showFirstRecordPreview(presetContext, variableNames, previewRow, getBlankBehaviorValue());
            } catch (previewError) {
                alert("The preview could not be generated.\n\n" + errorMessage(previewError), SCRIPT_NAME + " - Preview");
            }
        };

        backButton.onClick = function () {
            syncVisibleGrid();
            win.close(3); // Return to Stage 1 without losing records or current settings.
        };
        cancelButton.onClick = function () { win.close(0); };

        function finishGeneration(pdfMode) {
            syncVisibleGrid();
            if (records.length < 1) {
                alert(
                    "There are no records to generate. Add or import at least one row first.\n\n" +
                    "Finish does not save a permanent preset automatically. Use Export Preset... if you want to keep the setup and data for another session.",
                    SCRIPT_NAME + " - Finish"
                );
                return;
            }

            // Finish is deliberately session-only: generated artwork is added to the open
            // Illustrator document, but neither the AI file nor a persistent preset is saved.
            // Closing the document and choosing Don't Save discards this generation run.
            requestedPdfMode = pdfMode || "prompt";
            win.close(1);
        }

        exportPdfButton.onClick = function () {
            try {
                var rowsForPDF = getRecordCopy();
                if (!rowsForPDF.length) {
                    alert("Add or import at least one row before exporting a PDF.", SCRIPT_NAME + " - PDF Export");
                    return;
                }
                var pdfResult = exportRecordsAsSinglePDF(presetContext, variableNames, rowsForPDF, getBlankBehaviorValue());
                if (pdfResult && pdfResult.file) {
                    alert("PDF exported successfully. The data window is still open, and your records have been preserved.\n\nPages: " + pdfResult.count + "\nFile: " + pdfResult.file, SCRIPT_NAME + " - PDF Export");
                }
            } catch (independentPDFError) {
                alert("The PDF could not be exported. The data window remains open so you can continue editing.\n\n" + errorMessage(independentPDFError), SCRIPT_NAME + " - PDF Export");
            }
        };
        finishButton.onClick = function () { finishGeneration("prompt"); };

        function getBlankBehaviorValue() {
            return blankBehaviorDropdown.selection ? blankBehaviorDropdown.selection.text : "Clear mapped text";
        }

        function getRecordCopy() {
            syncVisibleGrid();
            var copied = [];
            var a;
            var b;
            for (a = 0; a < records.length; a++) {
                var rowCopy = [];
                for (b = 0; b < variableNames.length; b++) {
                    rowCopy.push(records[a][b] === undefined ? "" : String(records[a][b]));
                }
                copied.push(rowCopy);
            }
            return copied;
        }

        renderGrid();
        win.serializeData = serializeData;
        win.getBlankBehavior = getBlankBehaviorValue;
        win.getRecords = getRecordCopy;
        win.getPdfMode = function () { return requestedPdfMode; };
        return win;
    }

    function makeTextFrameLabel(frame, index) {
        var layerName = "(unknown layer)";
        var contents = "";
        try { if (frame.layer && frame.layer.name) { layerName = String(frame.layer.name); } } catch (ignoreLayerName) {}
        try { contents = String(frame.contents); } catch (ignoreFrameContents) {}
        if (contents === "") { contents = "(empty text frame)"; }
        contents = contents.replace(/\r/g, "\\r").replace(/\n/g, "\\n");
        if (contents.length > 42) { contents = contents.substring(0, 39) + "..."; }
        if (layerName.length > 26) { layerName = layerName.substring(0, 23) + "..."; }
        return "Text " + padNumber(index + 1, 2) + " | Layer: " + layerName + " | Text: " + contents;
    }

    function padNumber(n, width) {
        var result = String(n);
        while (result.length < width) { result = "0" + result; }
        return result;
    }

    function addLabeledInput(parent, labelText, defaultValue, characterCount) {
        var fieldGroup = parent.add("group");
        fieldGroup.orientation = "row";
        fieldGroup.alignChildren = ["center", "center"];
        fieldGroup.spacing = 4;
        fieldGroup.add("statictext", undefined, labelText);
        var field = fieldGroup.add("edittext", undefined, defaultValue);
        field.characters = characterCount;
        return field;
    }

    // ---------------------- Preset library / import ----------------------
    // Presets use a small tab-delimited format with URI-encoded fields. This avoids
    // requiring JSON support in older ExtendScript engines and safely preserves tabs,
    // line breaks, Unicode text, empty cells, and user-entered names.
    function presetEncode(value) {
        return encodeURIComponent(value === undefined || value === null ? "" : String(value));
    }

    function presetDecode(value) {
        return decodeURIComponent(String(value || ""));
    }

    function presetConfigKeys() {
        return ["layout", "artboardArrangement", "action", "columns", "rows", "gridGapX", "gridGapY", "horizontalGap",
            "verticalGap", "flattenStaticArtwork", "rasterResolution"];
    }

    function getTextFrameMetadata(frame, index) {
        var layerName = "";
        var contents = "";
        var bounds = "";
        try { contents = String(frame.contents); } catch (ignoreContents) {}
        try { layerName = String(frame.layer.name); } catch (ignoreLayer) {}
        try {
            var b = frame.geometricBounds;
            bounds = [roundToTenth(b[0]), roundToTenth(b[1]), roundToTenth(b[2]), roundToTenth(b[3])].join(",");
        } catch (ignoreBounds) {}
        return { index: index, contents: contents, layerName: layerName, bounds: bounds };
    }

    function roundToTenth(value) {
        return Math.round(Number(value) * 10) / 10;
    }

    function findPresetFrameMatch(savedMap, textFrames) {
        if (!savedMap || !textFrames || !textFrames.length) { return -1; }
        var savedContents = String(savedMap.sourceContents || "");
        var preferredIndex = Number(savedMap.frameIndex);
        if (isFinite(preferredIndex) && preferredIndex >= 0 && preferredIndex < textFrames.length) {
            var preferredMeta = getTextFrameMetadata(textFrames[preferredIndex], preferredIndex);
            if (preferredMeta.contents === savedContents && presetFrameSignatureMatches(savedMap, preferredMeta)) { return preferredIndex; }
        }
        var exactContents = [];
        var strongMatches = [];
        var i;
        for (i = 0; i < textFrames.length; i++) {
            var meta = getTextFrameMetadata(textFrames[i], i);
            if (meta.contents === savedContents) {
                exactContents.push(i);
                if (presetFrameSignatureMatches(savedMap, meta)) { strongMatches.push(i); }
            }
        }
        if (strongMatches.length === 1) { return strongMatches[0]; }
        // Content-only matching is allowed only when that content is unique in the selected group.
        return exactContents.length === 1 ? exactContents[0] : -1;
    }

    function presetFrameSignatureMatches(savedMap, currentMeta) {
        if (savedMap.layerName && currentMeta.layerName === savedMap.layerName) { return true; }
        if (savedMap.bounds && currentMeta.bounds === savedMap.bounds) { return true; }
        if (savedMap.normalizedBounds && currentMeta.bounds === savedMap.normalizedBounds) { return true; }
        return false;
    }

    function normalizeBoundsToArtboard(boundsString, artboardRect) {
        var parts = String(boundsString || "").split(",");
        if (parts.length !== 4 || !artboardRect || artboardRect.length !== 4) { return ""; }
        var b = [Number(parts[0]), Number(parts[1]), Number(parts[2]), Number(parts[3])];
        var h = Number(artboardRect[1]) - Number(artboardRect[3]);
        return [
            roundToTenth(b[0] - Number(artboardRect[0])),
            roundToTenth(b[1] + h - Number(artboardRect[1])),
            roundToTenth(b[2] - Number(artboardRect[0])),
            roundToTenth(b[3] + h - Number(artboardRect[1]))
        ].join(",");
    }

    function getPresetRootFolder() {
        var root = Folder.userData;
        if (!root || !root.exists) { root = Folder.myDocuments; }
        var productFolder = new Folder(root.fsName + "/Variable Data Studio");
        if (!productFolder.exists && !productFolder.create()) {
            throw new Error("Could not create the Variable Data Studio folder in your user data directory.");
        }
        return productFolder;
    }

    function getPresetLibraryFolder() {
        var productFolder = getPresetRootFolder();
        var libraryFolder = new Folder(productFolder.fsName + "/Library");
        if (!libraryFolder.exists && !libraryFolder.create()) {
            throw new Error("Could not create the preset Library folder.");
        }
        return libraryFolder;
    }

    function getLastPresetPathFile() {
        return new File(getPresetRootFolder().fsName + "/last-preset-path.txt");
    }

    function rememberLastPresetPath(presetFile) {
        if (!presetFile) { return; }
        var pathFile = getLastPresetPathFile();
        pathFile.encoding = "UTF-8";
        if (pathFile.open("w")) {
            pathFile.write(String(presetFile.fsName));
            pathFile.close();
        }
    }

    function getLastPresetFile() {
        try {
            var pathFile = getLastPresetPathFile();
            if (!pathFile.exists) { return null; }
            pathFile.encoding = "UTF-8";
            if (!pathFile.open("r")) { return null; }
            var pathText = trim(String(pathFile.read()).replace(/^\uFEFF/, ""));
            pathFile.close();
            if (!pathText) { return null; }
            var lastFile = new File(pathText);
            return lastFile.exists ? lastFile : null;
        } catch (ignoreLastPresetPath) { return null; }
    }

    function choosePresetFile(promptText) {
        var lastFile = getLastPresetFile();
        if (lastFile && confirm(
            "Load the last saved preset?\n\n" + lastFile.fsName + "\n\nChoose No to browse for another preset.",
            true,
            SCRIPT_NAME + " - Preset"
        )) {
            return lastFile;
        }
        return File.openDialog(promptText || "Select a Variable Data Studio preset (.vds or legacy .cbg)");
    }

    function resolvePresetTemplateFile(presetFile, preset, allowPrompt) {
        var candidate = null;
        try {
            if (preset && preset.templateBackupPath) {
                candidate = new File(preset.templateBackupPath);
                if (candidate.exists) { return candidate; }
            }
        } catch (ignoreStoredTemplatePath) {}
        try {
            if (presetFile) {
                var base = presetFile.name.replace(/\.(?:vds|cbg)$/i, "");
                candidate = new File(presetFile.path + "/" + base + "_Template.ai");
                if (candidate.exists) { return candidate; }
            }
        } catch (ignoreAdjacentTemplatePath) {}
        if (allowPrompt === false) { return null; }
        return File.openDialog("Select the Illustrator template file for this preset");
    }

    function findPresetTemplateGroup(documentRef, preset) {
        if (!documentRef) { return null; }
        var wantedName = String(preset && preset.templateName ? preset.templateName : "");
        var groups = [];
        var i;
        try {
            for (i = 0; i < documentRef.groupItems.length; i++) {
                groups.push(documentRef.groupItems[i]);
            }
        } catch (ignoreDocumentGroups) {}
        for (i = 0; i < groups.length; i++) {
            var groupName = "";
            try { groupName = String(groups[i].name || ""); } catch (ignoreGroupName) {}
            if (wantedName && groupName === wantedName) { return groups[i]; }
            if (!wantedName && groupName === "Source Template") { return groups[i]; }
        }
        // A standalone template backup should normally contain one top-level Group.
        var rootGroups = [];
        for (i = 0; i < groups.length; i++) {
            var isNested = false;
            try { isNested = groups[i].parent && groups[i].parent.typename === "GroupItem"; } catch (ignoreGroupParent) {}
            if (!isNested) { rootGroups.push(groups[i]); }
        }
        if (rootGroups.length === 1) { return rootGroups[0]; }
        if (groups.length === 1) { return groups[0]; }
        return null;
    }

    function sanitizePresetFileName(value) {
        var name = String(value || "Data Preset");
        name = name.replace(/[\\\/:*?"<>|\r\n]+/g, "_");
        name = name.replace(/^\s+|\s+$/g, "");
        if (!name) { name = "Data Preset"; }
        if (name.length > 100) { name = name.substring(0, 100); }
        return name;
    }

    function makePresetData(name, context, variableNames, records, blankBehavior, columnWidths, templateBackupPath) {
        var config = context.config;
        var sourceRect = context.sourceArtboardRect || [0, 0, 0, 0];
        var maps = [];
        var i;
        for (i = 0; i < variableNames.length; i++) {
            var frameIndex = context.frameIndices[i];
            var frame = context.templateFrames[frameIndex];
            var metadata = getTextFrameMetadata(frame, frameIndex);
            maps.push({
                variableName: String(variableNames[i]),
                frameIndex: frameIndex,
                sourceContents: metadata.contents,
                layerName: metadata.layerName,
                bounds: metadata.bounds,
                normalizedBounds: normalizeBoundsToArtboard(metadata.bounds, sourceRect)
            });
        }
        return {
            name: String(name),
            created: new Date().toString(),
            templateName: String(context.groupName || ""),
            templateBackupPath: String(templateBackupPath || ""),
            blankBehavior: String(blankBehavior || "Clear mapped text"),
            config: config,
            variables: variableNames.slice(0),
            maps: maps,
            records: cloneRecordRows(records, variableNames.length),
            columnWidths: cloneRecordRows([columnWidths || []], variableNames.length)[0],
            sourceArtboardRect: copyRect(sourceRect),
            sourceDocumentName: safeDocumentName(context.documentRef)
        };
    }

    function safeDocumentName(documentRef) {
        try { return String(documentRef.name); } catch (ignoreDocName) { return ""; }
    }

    function cloneRecordRows(records, columnCount) {
        var result = [];
        var i;
        var j;
        for (i = 0; i < records.length; i++) {
            var row = [];
            for (j = 0; j < columnCount; j++) {
                row.push(records[i] && records[i][j] !== undefined ? String(records[i][j]) : "");
            }
            result.push(row);
        }
        return result;
    }

    function serializePreset(preset) {
        var lines = ["VDS_PRESET\t1"];
        function add(tag, key, value) {
            lines.push(tag + "\t" + key + "\t" + presetEncode(value));
        }
        add("META", "name", preset.name);
        add("META", "created", preset.created);
        add("META", "templateName", preset.templateName);
        add("META", "templateBackupPath", preset.templateBackupPath);
        add("META", "blankBehavior", preset.blankBehavior);
        add("META", "sourceDocumentName", preset.sourceDocumentName);
        var rect = preset.sourceArtboardRect || [0, 0, 0, 0];
        add("META", "sourceArtboardRect", rect.join(","));
        var keys = presetConfigKeys();
        var i;
        for (i = 0; i < keys.length; i++) {
            var k = keys[i];
            add("CFG", k, preset.config[k]);
        }
        for (i = 0; i < preset.maps.length; i++) {
            var m = preset.maps[i];
            lines.push([
                "MAP", i, presetEncode(m.variableName), String(m.frameIndex),
                presetEncode(m.sourceContents), presetEncode(m.layerName), presetEncode(m.bounds), presetEncode(m.normalizedBounds)
            ].join("\t"));
        }
        for (i = 0; i < preset.variables.length; i++) {
            var storedWidth = preset.columnWidths && preset.columnWidths[i] !== undefined ? preset.columnWidths[i] : 0;
            if (Number(storedWidth) > 0) { lines.push("WIDTH\t" + i + "\t" + presetEncode(storedWidth)); }
        }
        for (i = 0; i < preset.records.length; i++) {
            var rowParts = ["ROW"];
            var j;
            for (j = 0; j < preset.variables.length; j++) {
                rowParts.push(presetEncode(preset.records[i][j] === undefined ? "" : preset.records[i][j]));
            }
            lines.push(rowParts.join("\t"));
        }
        return lines.join("\n");
    }

    function writePresetFile(file, preset) {
        file.encoding = "UTF-8";
        if (!file.open("w")) { throw new Error("Could not write preset file: " + file.fsName); }
        try {
            file.write(serializePreset(preset));
        } finally {
            file.close();
        }
    }

    function readPresetFile(file) {
        file.encoding = "UTF-8";
        if (!file.open("r")) { throw new Error("Could not open preset file: " + file.fsName); }
        var raw;
        try { raw = file.read(); } finally { file.close(); }
        raw = String(raw).replace(/^\uFEFF/, "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
        var lines = raw.split("\n");
        if (!lines.length || (lines[0] !== "VDS_PRESET\t1" && lines[0] !== "CBG_PRESET\t1")) {
            throw new Error("This file is not a supported Variable Data Studio preset.");
        }
        var preset = {
            name: "Data Preset", created: "", templateName: "", templateBackupPath: "",
            blankBehavior: "Clear mapped text", sourceDocumentName: "", sourceArtboardRect: [0, 0, 0, 0],
            config: {}, variables: [], maps: [], records: [], columnWidths: []
        };
        var i;
        for (i = 1; i < lines.length; i++) {
            if (lines[i] === "") { continue; }
            var parts = lines[i].split("\t");
            var type = parts[0];
            if (type === "META") {
                var value = presetDecode(parts.length > 2 ? parts.slice(2).join("\t") : "");
                if (parts[1] === "name") { preset.name = value; }
                else if (parts[1] === "created") { preset.created = value; }
                else if (parts[1] === "templateName") { preset.templateName = value; }
                else if (parts[1] === "templateBackupPath") { preset.templateBackupPath = value; }
                else if (parts[1] === "blankBehavior") { preset.blankBehavior = value; }
                else if (parts[1] === "sourceDocumentName") { preset.sourceDocumentName = value; }
                else if (parts[1] === "sourceArtboardRect") {
                    var rectParts = value.split(",");
                    if (rectParts.length === 4) { preset.sourceArtboardRect = [Number(rectParts[0]), Number(rectParts[1]), Number(rectParts[2]), Number(rectParts[3])]; }
                }
            } else if (type === "CFG") {
                var cfgValue = presetDecode(parts.length > 2 ? parts.slice(2).join("\t") : "");
                var key = parts[1];
                if (key === "columns" || key === "rows" || key === "rasterResolution") { cfgValue = Number(cfgValue); }
                else if (key === "gridGapX" || key === "gridGapY" || key === "horizontalGap" || key === "verticalGap") { cfgValue = Number(cfgValue); }
                else if (key === "flattenStaticArtwork") { cfgValue = (cfgValue === "true"); }
                preset.config[key] = cfgValue;
            } else if (type === "MAP") {
                var mapIndex = Number(parts[1]);
                preset.maps[mapIndex] = {
                    variableName: presetDecode(parts[2] || ""),
                    frameIndex: Number(parts[3]),
                    sourceContents: presetDecode(parts[4] || ""),
                    layerName: presetDecode(parts[5] || ""),
                    bounds: presetDecode(parts[6] || ""),
                    normalizedBounds: presetDecode(parts[7] || "")
                };
                preset.variables[mapIndex] = preset.maps[mapIndex].variableName;
            } else if (type === "WIDTH") {
                var widthIndex = Number(parts[1]);
                preset.columnWidths[widthIndex] = Number(presetDecode(parts[2] || "0")) || 0;
            } else if (type === "ROW") {
                var row = [];
                var j;
                for (j = 1; j < parts.length; j++) { row.push(presetDecode(parts[j])); }
                while (row.length < preset.variables.length) { row.push(""); }
                preset.records.push(row);
            }
        }
        if (!preset.variables.length || !preset.maps.length) { throw new Error("The preset does not contain variable mappings. It may store settings, but a mapped text frame is required to generate outputs."); }
        // Empty records are valid: a preset can store stage-one configuration before any data is entered.
        // v1.7 and earlier did not store the artboard arrangement; preserve backward compatibility.
        if (preset.config.artboardArrangement === undefined) { preset.config.artboardArrangement = "Grid"; }
        var keyNames = presetConfigKeys();
        for (i = 0; i < keyNames.length; i++) {
            if (preset.config[keyNames[i]] === undefined) { throw new Error("The preset is missing setting '" + keyNames[i] + "'."); }
        }
        return preset;
    }

    function savePresetPackage(presetFile, presetName, context, variableNames, records, blankBehavior, columnWidths, saveBesideTemplate) {
        var warning = "";
        var templateBackupPath = "";
        if (saveBesideTemplate) {
            var basePath = presetFile.fsName.replace(/\.(?:vds|cbg)$/i, "");
            var backupFile = new File(basePath + "_Template.ai");
            try {
                saveTemplateBackup(context.documentRef, context.templateGroup, context.sourceArtboardRect, backupFile);
                templateBackupPath = backupFile.fsName;
            } catch (backupError) {
                warning = "The preset was saved, but a standalone template copy could not be created: " + errorMessage(backupError) + ". You may need to reassign text frames when importing this preset.";
            }
        }
        var preset = makePresetData(presetName, context, variableNames, records, blankBehavior, columnWidths, templateBackupPath);
        writePresetFile(presetFile, preset);
        return { templateBackupPath: templateBackupPath, warning: warning };
    }

    function saveTemplateBackup(sourceDocument, sourceGroup, sourceArtboardRect, backupFile) {
        if (!sourceDocument || !sourceGroup) { throw new Error("The source template is not available."); }
        var artboardWidth = Number(sourceArtboardRect[2]) - Number(sourceArtboardRect[0]);
        var artboardHeight = Number(sourceArtboardRect[1]) - Number(sourceArtboardRect[3]);
        if (!(artboardWidth > 0) || !(artboardHeight > 0)) { throw new Error("The source artboard has invalid dimensions."); }
        var tempPath = backupFile.fsName.replace(/\.ai$/i, "") + "_tmp_" + (new Date().getTime()) + ".ai";
        var tempFile = new File(tempPath);
        var backupDocument = null;
        try {
            backupDocument = app.documents.add(sourceDocument.documentColorSpace, artboardWidth, artboardHeight);
            backupDocument.artboards[0].artboardRect = [0, artboardHeight, artboardWidth, 0];
            app.activeDocument = sourceDocument;
            var duplicate = sourceGroup.duplicate(backupDocument.layers[0], ElementPlacement.PLACEATBEGINNING);
            duplicate.translate(-sourceArtboardRect[0], artboardHeight - sourceArtboardRect[1]);
            try { duplicate.name = String(sourceGroup.name || "Source Template"); } catch (ignoreBackupGroupName) {}
            app.activeDocument = backupDocument;
            var options = new IllustratorSaveOptions();
            options.pdfCompatible = true;
            backupDocument.saveAs(tempFile, options);
            backupDocument.close(SaveOptions.DONOTSAVECHANGES);
            backupDocument = null;
            if (backupFile.exists && !backupFile.remove()) { throw new Error("Could not replace the previous template backup."); }
            if (!tempFile.copy(backupFile.fsName)) { throw new Error("Could not move the saved template copy into the preset folder."); }
            try { tempFile.remove(); } catch (ignoreTempDelete) {}
            app.activeDocument = sourceDocument;
            return backupFile.fsName;
        } catch (saveTemplateError) {
            try { if (backupDocument) { backupDocument.close(SaveOptions.DONOTSAVECHANGES); } } catch (ignoreBackupClose) {}
            try { if (tempFile.exists) { tempFile.remove(); } } catch (ignoreTempDeleteError) {}
            try { app.activeDocument = sourceDocument; } catch (ignoreRestoreDocument) {}
            throw saveTemplateError;
        }
    }

    function readConfiguration(win) {
        var config = readConfigurationFromControls(
            win.layoutDropdown,
            win.actionDropdown,
            win.columnsInput,
            win.rowsInput,
            win.horizontalGridSpacingInput,
            win.verticalGridSpacingInput,
            win.horizontalSpacingInput,
            win.verticalSpacingInput,
            win.flattenStaticCheckbox,
            win.rasterResolutionDropdown,
            win.artboardArrangementDropdown
        );
        var maps = win.readMappings();
        config.variables = maps.variables;
        config.frameIndices = maps.frameIndices;
        return config;
    }

    function readConfigurationFromControls(layoutDropdown, actionDropdown, columnsInput, rowsInput,
        horizontalGridSpacingInput, verticalGridSpacingInput, horizontalSpacingInput, verticalSpacingInput,
        flattenStaticCheckbox, rasterResolutionDropdown, artboardArrangementDropdown) {
        var resolutionText = rasterResolutionDropdown && rasterResolutionDropdown.selection ? rasterResolutionDropdown.selection.text : "300 ppi";
        return {
            layout: layoutDropdown.selection.text,
            artboardArrangement: artboardArrangementDropdown && artboardArrangementDropdown.selection ? artboardArrangementDropdown.selection.text : "Grid",
            action: actionDropdown.selection.text,
            columns: parsePositiveInteger(columnsInput.text, "Columns"),
            rows: parsePositiveInteger(rowsInput.text, "Rows"),
            gridGapX: parseNonNegativeNumber(horizontalGridSpacingInput.text, "Horizontal gap"),
            gridGapY: parseNonNegativeNumber(verticalGridSpacingInput.text, "Vertical gap"),
            horizontalGap: parseNonNegativeNumber(horizontalSpacingInput.text, "Horizontal spacing"),
            verticalGap: parseNonNegativeNumber(verticalSpacingInput.text, "Vertical spacing"),
            flattenStaticArtwork: flattenStaticCheckbox ? Boolean(flattenStaticCheckbox.value) : false,
            rasterResolution: parseInt(resolutionText, 10) || 300
        };
    }

    function makeDataRowsFromGrid(gridRecords, variableNames) {
        var rows = [];
        var i;
        var j;
        if (!gridRecords || !gridRecords.length) {
            throw new Error("No records were supplied. Add rows or import data, then press Finish.");
        }
        for (i = 0; i < gridRecords.length; i++) {
            var record = {};
            for (j = 0; j < variableNames.length; j++) {
                record[variableNames[j]] = gridRecords[i] && gridRecords[i][j] !== undefined
                    ? String(gridRecords[i][j]) : "";
            }
            rows.push(record);
        }
        return { rows: rows };
    }

    // ----------------------- Legacy TSV parsing (kept for compatibility) -----------------------
    function parseRecordTable(rawText, variableNames) {
        var normalizedText = String(rawText).replace(/^\uFEFF/, "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
        var lines = normalizedText.split("\n");
        var headerLineIndex = -1;
        var headerCells;
        var variableColumnIndices = {};
        var i;
        var j;

        for (i = 0; i < lines.length; i++) {
            if (trim(lines[i]) !== "") {
                headerLineIndex = i;
                break;
            }
        }
        if (headerLineIndex < 0) {
            throw new Error("The data grid is empty. Add at least one record or import data from Excel.");
        }

        headerCells = splitTsvLine(lines[headerLineIndex]);
        if (headerCells.length === 1 && trim(headerCells[0]) === "") {
            throw new Error("The variable header row is empty.");
        }

        for (i = 0; i < headerCells.length; i++) {
            var headerText = trim(headerCells[i]);
            var key = normalizeVariableName(headerText);
            if (!key) {
                throw new Error("A variable header is empty. Check the Variable Mapping stage.");
            }
            if (variableColumnIndices[key] !== undefined) {
                throw new Error("The header row contains a duplicate variable name: " + headerText);
            }
            variableColumnIndices[key] = i;
        }

        for (i = 0; i < variableNames.length; i++) {
            var requiredKey = normalizeVariableName(variableNames[i]);
            if (variableColumnIndices[requiredKey] === undefined) {
                throw new Error("The data grid is missing the configured variable '" + variableNames[i] + "'.");
            }
        }
        if (headerCells.length !== variableNames.length) {
            throw new Error("The data grid must contain exactly " + variableNames.length + " variable column(s).");
        }

        var rows = [];
        for (i = headerLineIndex + 1; i < lines.length; i++) {
            var line = lines[i];
            // Preserve empty lines because a one-variable record with a blank value serializes as an empty line.
            var cells = splitTsvLine(line);
            while (cells.length < headerCells.length) { cells.push(""); }
            if (cells.length > headerCells.length) {
                for (j = headerCells.length; j < cells.length; j++) {
                    if (trim(cells[j]) !== "") {
                        throw new Error("Record row " + (i - headerLineIndex) + " has more columns than the variable header row.");
                    }
                }
            }
            var record = {};
            for (j = 0; j < variableNames.length; j++) {
                var varName = variableNames[j];
                var cellValue = cells[variableColumnIndices[normalizeVariableName(varName)]];
                record[varName] = cellValue === undefined ? "" : String(cellValue);
            }
            rows.push(record);
        }

        if (rows.length === 0) {
            throw new Error("No records were found. Add a row or import data from Excel.");
        }
        return { rows: rows };
    }

    function splitTsvLine(line) {
        // Tabs inside a single value are not supported; standard Excel TSV paste is supported.
        return String(line).split("\t");
    }

    function isSerialHeader(value) {
        var s = trim(value).toLowerCase().replace(/\s+/g, " ");
        return s === "#" || s === "no" || s === "no." || s === "serial" ||
            s === "serial no" || s === "serial no." || s === "serial number" ||
            s === "record #" || s === "record no" || s === "record number" ||
            s === "رقم" || s === "الرقم" || s === "رقم تسلسلي" || s === "الرقم التسلسلي" || s === "التسلسل";
    }

    function normalizeVariableName(value) {
        var s = trim(String(value)).toLowerCase().replace(/\s+/g, " ");
        return s ? s : null;
    }

    // -------------------- Generation routines --------------------
    function generateOnCanvas(documentRef, sourceGroup, variableNames, frameIndices, rows, sourceBounds, config, created) {
        var output = [];
        var width = sourceBounds[2] - sourceBounds[0];
        var height = sourceBounds[1] - sourceBounds[3];
        var hGap = 0;
        var vGap = 0;
        var baseLeft = sourceBounds[0];
        var baseTop = sourceBounds[1];
        var i;

        if (config.layout === "Grid") {
            hGap = cmToPoints(config.gridGapX);
            vGap = cmToPoints(config.gridGapY);
        } else if (config.layout === "Horizontal") {
            hGap = cmToPoints(config.horizontalGap);
        } else if (config.layout === "Vertical") {
            vGap = cmToPoints(config.verticalGap);
        }

        if (config.action === "Keep Original") {
            if (config.layout === "Grid" || config.layout === "Horizontal") { baseLeft = sourceBounds[2] + hGap; }
            else if (config.layout === "Vertical") { baseTop = sourceBounds[3] - vGap; }
        }

        for (i = 0; i < rows.length; i++) {
            if (config.action === "Replace Original" && i === 0 && !config.flattenStaticArtwork) {
                output[i] = sourceGroup;
                continue;
            }

            var clone;
            var preReplacementBounds;
            if (config.flattenStaticArtwork) {
                clone = createRasterizedOutput(documentRef, sourceGroup, variableNames, frameIndices, rows[i], config, sourceBounds, created);
                preReplacementBounds = copyRect(sourceBounds);
            } else {
                clone = sourceGroup.duplicate();
                created.push(clone);
                preReplacementBounds = getArtworkBounds(clone);
                replacePlaceholders(clone, variableNames, frameIndices, rows[i], config.blankBehavior, i + 1);
            }

            var targetLeft = sourceBounds[0];
            var targetTop = sourceBounds[1];
            if (config.layout === "Grid") {
                targetLeft = baseLeft + (i % config.columns) * (width + hGap);
                targetTop = baseTop - Math.floor(i / config.columns) * (height + vGap);
            } else if (config.layout === "Horizontal") {
                targetLeft = baseLeft + i * (width + hGap);
                targetTop = baseTop;
            } else if (config.layout === "Vertical") {
                targetLeft = baseLeft;
                targetTop = baseTop - i * (height + vGap);
            } else if (config.layout === "Same Position") {
                targetLeft = sourceBounds[0];
                targetTop = sourceBounds[1];
            }

            translateUsingBounds(clone, preReplacementBounds, targetLeft, targetTop);
            try { clone.name = makeOutputName(i + 1); } catch (ignoreCloneName) {}
            output[i] = clone;
        }
        return output;
    }

    function generateOnSeparateArtboards(documentRef, sourceGroup, variableNames, frameIndices, rows, sourceBounds,
        existingArtboardCount, sourceArtboardRect, config, createdItemsRef, createdArtboardIndicesRef) {
        var output = [];
        var boardWidth = sourceArtboardRect[2] - sourceArtboardRect[0];
        var boardHeight = sourceArtboardRect[1] - sourceArtboardRect[3];
        var boardGapX = cmToPoints(config.gridGapX);
        var boardGapY = cmToPoints(config.gridGapY);
        var maxRight = -Infinity;
        var maxTop = -Infinity;
        var i;

        if (!(boardWidth > 0) || !(boardHeight > 0)) {
            throw new Error("The source artboard has an invalid width or height.");
        }
        for (i = 0; i < existingArtboardCount; i++) {
            try {
                var rect = documentRef.artboards[i].artboardRect;
                if (rect[2] > maxRight) { maxRight = rect[2]; }
                if (rect[1] > maxTop) { maxTop = rect[1]; }
            } catch (readBoardError) {
                throw new Error("Could not read existing artboard #" + (i + 1) + ": " + errorMessage(readBoardError));
            }
        }
        if (!isFinite(maxRight) || !isFinite(maxTop)) { throw new Error("Could not determine existing artboard positions."); }

        var arrangement = config.artboardArrangement || "Grid";
        var columns = config.columns;
        if (arrangement === "Horizontal only") { columns = rows.length; }
        else if (arrangement === "Vertical only") { columns = 1; }
        if (!(columns >= 1)) { throw new Error("The artboard column count is invalid."); }

        var targetRects = [];
        // Create and record each artboard immediately. If Illustrator rejects a coordinate or
        // returns a stale collection element, the error identifies the exact record being placed.
        for (i = 0; i < rows.length; i++) {
            var col = i % columns;
            var row = Math.floor(i / columns);
            var left = maxRight + boardGapX + col * (boardWidth + boardGapX);
            var top = maxTop - row * (boardHeight + boardGapY);
            var artboardRect = [left, top, left + boardWidth, top - boardHeight];
            var newArtboard = null;
            try {
                newArtboard = documentRef.artboards.add(artboardRect);
                createdArtboardIndicesRef.push(documentRef.artboards.length - 1);
            } catch (addBoardError) {
                throw new Error("Could not create the artboard for record #" + (i + 1) + ": " + errorMessage(addBoardError));
            }
            try { newArtboard.name = makeOutputName(i + 1); } catch (ignoreArtboardName) {}
            targetRects.push(artboardRect);
        }

        for (i = 0; i < rows.length; i++) {
            var recordNumber = i + 1;
            if (config.action === "Replace Original" && i === 0 && !config.flattenStaticArtwork) {
                // Stage record #1 on the first new artboard; the source Group is updated only
                // after all generated outputs have succeeded.
                var staged = null;
                try {
                    staged = sourceGroup.duplicate();
                    createdItemsRef.push(staged);
                } catch (duplicateStagedError) {
                    throw new Error("Could not duplicate the template for record #" + recordNumber + ": " + errorMessage(duplicateStagedError));
                }
                try {
                    replacePlaceholders(staged, variableNames, frameIndices, rows[i], config.blankBehavior, recordNumber);
                } catch (replaceStagedError) {
                    throw new Error("Could not replace mapped text for record #" + recordNumber + ": " + errorMessage(replaceStagedError));
                }
                var stagedDeltaX = targetRects[i][0] - sourceArtboardRect[0];
                var stagedDeltaY = targetRects[i][1] - sourceArtboardRect[1];
                try { staged.translate(stagedDeltaX, stagedDeltaY); }
                catch (translateStagedError) { throw new Error("Could not position record #" + recordNumber + ": " + errorMessage(translateStagedError)); }
                try { staged.name = makeOutputName(recordNumber); } catch (ignoreStagedName) {}
                output.push(staged);
                continue;
            }

            var clone = null;
            if (config.flattenStaticArtwork) {
                try {
                    clone = createRasterizedOutput(documentRef, sourceGroup, variableNames, frameIndices, rows[i], config, sourceArtboardRect, createdItemsRef);
                } catch (flattenOutputError) {
                    throw new Error("Could not prepare flattened artwork for record #" + recordNumber + ": " + errorMessage(flattenOutputError));
                }
            } else {
                try {
                    clone = sourceGroup.duplicate();
                    createdItemsRef.push(clone);
                } catch (duplicateOutputError) {
                    throw new Error("Could not duplicate the template for record #" + recordNumber + ": " + errorMessage(duplicateOutputError));
                }
                try {
                    replacePlaceholders(clone, variableNames, frameIndices, rows[i], config.blankBehavior, recordNumber);
                } catch (replaceOutputError) {
                    throw new Error("Could not replace mapped text for record #" + recordNumber + ": " + errorMessage(replaceOutputError));
                }
            }
            var dx = targetRects[i][0] - sourceArtboardRect[0];
            var dy = targetRects[i][1] - sourceArtboardRect[1];
            try { clone.translate(dx, dy); }
            catch (translateOutputError) { throw new Error("Could not position record #" + recordNumber + ": " + errorMessage(translateOutputError)); }
            try { clone.name = makeOutputName(recordNumber); } catch (ignoreArtName) {}
            output.push(clone);
        }
        return output;
    }

    function createRasterizedOutput(documentRef, sourceGroup, variableNames, frameIndices, record, config, sourceBounds, created) {
        var staticClone = sourceGroup.duplicate();
        created.push(staticClone);
        var staticFrames = collectTextFrames(staticClone);
        var removeRefs = [];
        var i;
        for (i = 0; i < frameIndices.length; i++) {
            var mappedIndex = frameIndices[i];
            if (mappedIndex < 0 || mappedIndex >= staticFrames.length) {
                throw new Error("A mapped text frame could not be found while flattening the static artwork.");
            }
            addUniqueReference(removeRefs, staticFrames[mappedIndex]);
        }
        for (i = removeRefs.length - 1; i >= 0; i--) {
            try { removeRefs[i].remove(); }
            catch (removeFrameError) { throw new Error("Could not isolate the editable variable text frames for rasterization: " + errorMessage(removeFrameError)); }
        }

        var rasterOptions = new RasterizeOptions();
        rasterOptions.resolution = config.rasterResolution || 300;
        rasterOptions.transparency = true;
        rasterOptions.clippingMask = false;
        rasterOptions.padding = 0;
        try { rasterOptions.antiAliasingMethod = AntiAliasingMethod.ARTOPTIMIZED; } catch (ignoreAntiAlias) {}
        try { rasterOptions.includeLayers = true; } catch (ignoreIncludeLayers) {}
        var raster = documentRef.rasterize(staticClone, copyRect(sourceBounds), rasterOptions);
        created.push(raster);
        try { raster.name = "Static_Artwork"; } catch (ignoreRasterName) {}

        var parentContainer = sourceGroup.parent;
        if (!parentContainer || !parentContainer.groupItems) {
            throw new Error("Could not create an output group beside the selected template.");
        }
        var wrapper = parentContainer.groupItems.add();
        created.push(wrapper);
        try { wrapper.name = "Flattened_Output"; } catch (ignoreWrapperName) {}
        raster.move(wrapper, ElementPlacement.PLACEATEND);

        var liveTextGroup = wrapper.groupItems.add();
        try { liveTextGroup.name = "Editable_Variables"; } catch (ignoreLiveTextName) {}
        var sourceFrames = collectTextFrames(sourceGroup);
        var liveFrames = [];
        for (i = 0; i < frameIndices.length; i++) {
            var sourceIndex = frameIndices[i];
            if (sourceIndex < 0 || sourceIndex >= sourceFrames.length) {
                throw new Error("A mapped text frame could not be found while creating editable text overlays.");
            }
            var liveFrame = sourceFrames[sourceIndex].duplicate(liveTextGroup, ElementPlacement.PLACEATEND);
            liveFrames.push(liveFrame);
            var value = record[variableNames[i]] === undefined ? "" : String(record[variableNames[i]]);
            if (!(config.blankBehavior === "Keep template text" && value === "")) {
                replaceContentsPreservingStyle(liveFrame, value);
            }
        }
        return wrapper;
    }

    function replacePlaceholders(group, variableNames, frameIndices, record, blankBehavior, recordNumber) {
        var frames = collectTextFrames(group);
        var i;
        if (!variableNames || !frameIndices || variableNames.length !== frameIndices.length) {
            throw new Error("The variable mapping is invalid.");
        }
        for (i = 0; i < variableNames.length; i++) {
            var frameIndex = frameIndices[i];
            if (frameIndex < 0 || frameIndex >= frames.length || !frames[frameIndex]) {
                throw new Error("The assigned text frame for " + variableNames[i] + " could not be found in a output copy.");
            }
            var value = record[variableNames[i]] === undefined ? "" : String(record[variableNames[i]]);
            if (blankBehavior === "Keep template text" && value === "") { continue; }
            try {
                replaceContentsPreservingStyle(frames[frameIndex], value);
            } catch (replaceTextError) {
                throw new Error("Record #" + (recordNumber || "?") + ", variable '" + variableNames[i] + "': " + errorMessage(replaceTextError));
            }
        }
    }

    // Replace text while restoring original character/paragraph styling by position.
    // IMPORTANT: strokeColor and strokeWeight are intentionally excluded from both capture and replay.
    // Illustrator may report default/stale stroke values for unstroked text; writing them can add a black
    // outline. The duplicated frame's existing stroke appearance is left untouched.
    // Per-character snapshots preserve mixed formatting as closely as ExtendScript allows.
    function replaceContentsPreservingStyle(frame, newContents) {
        // Capture style before changing the string. A duplicated Illustrator TextFrame already
        // carries its original appearance, so preserving that live frame is the primary safeguard.
        // Reapplying individual style runs is best-effort: some Illustrator builds expose stale
        // Characters/Paragraphs collection entries after the contents length changes. Those
        // collection errors must never abort the entire batch or trigger a document rollback.
        var characterStyles = [];
        var paragraphStyles = [];
        try { characterStyles = captureCharacterStyleRuns(frame); } catch (ignoreCaptureCharacterRuns) {}
        try { paragraphStyles = captureParagraphStyleRuns(frame); } catch (ignoreCaptureParagraphRuns) {}
        var fallbackCharacterStyle = characterStyles.length ? characterStyles[0] : null;

        try { frame.contents = String(newContents); }
        catch (setContentsError) { throw new Error("Could not set the text contents: " + errorMessage(setContentsError)); }

        var range = null;
        try { range = frame.textRange; } catch (getRangeError) { return; }
        if (!range) { return; }

        var chars = null;
        try { chars = range.characters; } catch (getCharactersError) { chars = null; }
        if (chars) {
            var charCount = 0;
            try { charCount = chars.length; } catch (getCharacterCountError) { charCount = 0; }
            var i;
            for (i = 0; i < charCount; i++) {
                var character = null;
                var characterAttributes = null;
                try { character = chars[i]; } catch (getCharacterError) { continue; }
                if (!character) { continue; }
                try { characterAttributes = character.characterAttributes; } catch (getCharacterAttributesError) { continue; }
                var style = characterStyles.length ? characterStyles[Math.min(i, characterStyles.length - 1)] : fallbackCharacterStyle;
                if (style && characterAttributes) {
                    try { applyAttributeSnapshot(characterAttributes, style.attributes); } catch (applyCharacterStyleError) {}
                }
            }
        } else if (fallbackCharacterStyle) {
            // Empty text frames have no Characters collection. The original frame retains its
            // appearance; do not force synthetic attributes into an empty range.
        }

        var newParagraphs = null;
        try { newParagraphs = range.paragraphs; } catch (getParagraphsError) { newParagraphs = null; }
        if (newParagraphs && paragraphStyles.length) {
            var paragraphCount = 0;
            try { paragraphCount = newParagraphs.length; } catch (getParagraphCountError) { paragraphCount = 0; }
            for (i = 0; i < paragraphCount; i++) {
                var paragraph = null;
                var paragraphAttributes = null;
                try { paragraph = newParagraphs[i]; } catch (getParagraphError) { continue; }
                if (!paragraph) { continue; }
                try { paragraphAttributes = paragraph.paragraphAttributes; } catch (getParagraphAttributesError) { continue; }
                try {
                    applyAttributeSnapshot(
                        paragraphAttributes,
                        paragraphStyles[Math.min(i, paragraphStyles.length - 1)]
                    );
                } catch (applyParagraphStyleError) {}
            }
        }
        // strokeColor and strokeWeight remain intentionally excluded from all snapshots/replay.
    }

    function captureCharacterStyleRuns(frame) {
        var result = [];
        var chars = null;
        var i;
        try { chars = frame.textRange.characters; } catch (ignoreSourceCharacters) {}
        if (chars && chars.length) {
            for (i = 0; i < chars.length; i++) {
                var attr = null;
                try { attr = chars[i].characterAttributes; } catch (ignoreCharacterAttributes) {}
                result.push({
                    attributes: captureAttributeSnapshot(attr, CHARACTER_ATTRIBUTE_NAMES)
                });
            }
        }
        if (!result.length) {
            var rangeAttributes = null;
            try { rangeAttributes = frame.textRange.characterAttributes; } catch (ignoreRangeCharacterAttributes) {}
            result.push({
                attributes: captureAttributeSnapshot(rangeAttributes, CHARACTER_ATTRIBUTE_NAMES)
            });
        }
        return result;
    }

    function captureParagraphStyleRuns(frame) {
        var result = [];
        var paragraphs = null;
        var i;
        try { paragraphs = frame.textRange.paragraphs; } catch (ignoreSourceParagraphs) {}
        if (paragraphs && paragraphs.length) {
            for (i = 0; i < paragraphs.length; i++) {
                var attrs = null;
                try { attrs = paragraphs[i].paragraphAttributes; } catch (ignoreParagraphAttributes) {}
                result.push(captureAttributeSnapshot(attrs, PARAGRAPH_ATTRIBUTE_NAMES));
            }
        }
        if (!result.length) {
            var source = null;
            try { source = frame.textRange.paragraphAttributes; } catch (ignoreRangeParagraphAttributes) {}
            result.push(captureAttributeSnapshot(source, PARAGRAPH_ATTRIBUTE_NAMES));
        }
        return result;
    }

    function captureAttributeSnapshot(source, propertyNames) {
        var snapshot = [];
        var i;
        if (!source) { return snapshot; }
        for (i = 0; i < propertyNames.length; i++) {
            try {
                var value = source[propertyNames[i]];
                if (value !== undefined && value !== null) {
                    snapshot.push({ name: propertyNames[i], value: value });
                }
            } catch (ignoreReadAttribute) {}
        }
        return snapshot;
    }

    function applyAttributeSnapshot(target, snapshot) {
        var i;
        if (!target || !snapshot) { return; }
        for (i = 0; i < snapshot.length; i++) {
            try { target[snapshot[i].name] = snapshot[i].value; } catch (ignoreWriteAttribute) {}
        }
    }

    // --------------------- Text-frame discovery ---------------------
    function getSelectedTemplate(documentRef) {
        try {
            if (!documentRef.selection || documentRef.selection.length !== 1) {
                return null;
            }
            if (documentRef.selection[0].typename !== "GroupItem") {
                return null;
            }
            return documentRef.selection[0];
        } catch (selectionError) {
            return null;
        }
    }

    function collectTextFrames(group) {
        var found = [];
        collectTextFramesRecursive(group, found);
        return found;
    }

    function collectTextFramesRecursive(group, found) {
        var i;
        try {
            var textFrames = group.textFrames;
            for (i = 0; i < textFrames.length; i++) {
                addUniqueReference(found, textFrames[i]);
            }
        } catch (ignoreTextCollection) {}

        try {
            var nestedGroups = group.groupItems;
            for (i = 0; i < nestedGroups.length; i++) {
                if (isDirectChild(nestedGroups[i], group)) {
                    collectTextFramesRecursive(nestedGroups[i], found);
                }
            }
        } catch (ignoreNestedGroups) {}
    }

    function snapshotTextFrames(frames) {
        var snapshot = [];
        var i;
        for (i = 0; i < frames.length; i++) {
            snapshot.push({ frame: frames[i], contents: String(frames[i].contents) });
        }
        return snapshot;
    }

    function restoreTextSnapshot(snapshot) {
        var i;
        for (i = 0; i < snapshot.length; i++) {
            try { snapshot[i].frame.contents = snapshot[i].contents; } catch (ignoreRestoreText) {}
        }
    }

    function findContainingArtboardIndex(documentRef, itemBounds, preferredIndex) {
        var bestIndex = preferredIndex;
        var bestArea = -1;
        var itemWidth = Math.max(0, itemBounds[2] - itemBounds[0]);
        var itemHeight = Math.max(0, itemBounds[1] - itemBounds[3]);
        var itemCenterX = (itemBounds[0] + itemBounds[2]) / 2;
        var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;
        var i;
        for (i = 0; i < documentRef.artboards.length; i++) {
            var rect = documentRef.artboards[i].artboardRect;
            var left = Math.max(itemBounds[0], rect[0]);
            var top = Math.min(itemBounds[1], rect[1]);
            var right = Math.min(itemBounds[2], rect[2]);
            var bottom = Math.max(itemBounds[3], rect[3]);
            var overlapWidth = Math.max(0, right - left);
            var overlapHeight = Math.max(0, top - bottom);
            var overlap = overlapWidth * overlapHeight;
            var containsCenter = itemCenterX >= rect[0] && itemCenterX <= rect[2] && itemCenterY <= rect[1] && itemCenterY >= rect[3];
            if (containsCenter) { overlap += Math.max(1, itemWidth * itemHeight) * 1000; }
            if (i === preferredIndex && overlap > 0) { overlap += 0.01; }
            if (overlap > bestArea) { bestArea = overlap; bestIndex = i; }
        }
        return bestIndex;
    }

    // -------------------- Position and artboard helpers --------------------
    // The outermost clipping group is the most reliable definition of the finished
    // output edge. Do not use masks from unrelated photos or icons as the artboard.
    // If the selected Group is not clipped, a unique shallow clipped child may define the
    // output boundary. Ambiguous inner masks fall back to the full Group's visible bounds.
    function getArtworkBounds(item) {
        var clippedGroup = null;
        try {
            if (item.typename === "GroupItem" && item.clipped === true) {
                clippedGroup = item;
            }
        } catch (ignoreRootClipFlag) {}

        if (!clippedGroup) {
            var clippedGroups = [];
            collectClippedGroups(item, 1, clippedGroups);
            if (clippedGroups.length > 0) {
                var shallowestDepth = clippedGroups[0].depth;
                var shallowest = [];
                var i;
                for (i = 0; i < clippedGroups.length; i++) {
                    if (clippedGroups[i].depth < shallowestDepth) {
                        shallowestDepth = clippedGroups[i].depth;
                        shallowest = [];
                    }
                    if (clippedGroups[i].depth === shallowestDepth) {
                        shallowest.push(clippedGroups[i].group);
                    }
                }
                if (shallowest.length === 1) { clippedGroup = shallowest[0]; }
            }
        }

        if (clippedGroup) {
            var clippingPaths = [];
            collectDirectClippingPaths(clippedGroup, clippingPaths);
            var bestBounds = null;
            var bestArea = -1;
            var j;
            for (j = 0; j < clippingPaths.length; j++) {
                var maskBounds = getGeometricBounds(clippingPaths[j]);
                var maskWidth = Math.max(0, maskBounds[2] - maskBounds[0]);
                var maskHeight = Math.max(0, maskBounds[1] - maskBounds[3]);
                var area = maskWidth * maskHeight;
                if (area > bestArea) {
                    bestArea = area;
                    bestBounds = maskBounds;
                }
            }
            if (bestBounds) { return bestBounds; }
        }
        return getBounds(item);
    }

    function collectClippedGroups(container, depth, found) {
        var i;
        try {
            var groups = container.groupItems;
            for (i = 0; i < groups.length; i++) {
                var nested = groups[i];
                if (!isDirectChild(nested, container)) { continue; }
                var isClipped = false;
                try { isClipped = nested.clipped === true; } catch (ignoreNestedClipFlag) {}
                if (isClipped) { addClippedGroupCandidate(found, nested, depth); }
                collectClippedGroups(nested, depth + 1, found);
            }
        } catch (ignoreNestedClippedGroups) {}
    }

    function addClippedGroupCandidate(found, group, depth) {
        var i;
        for (i = 0; i < found.length; i++) {
            if (found[i].group === group) {
                if (depth < found[i].depth) { found[i].depth = depth; }
                return;
            }
        }
        found.push({ group: group, depth: depth });
    }

    function isDirectChild(item, expectedParent) {
        try {
            if (item.parent !== undefined && item.parent !== null) {
                return item.parent === expectedParent;
            }
        } catch (ignoreParentRead) {}
        return true;
    }

    function collectDirectClippingPaths(container, found) {
        var i;
        try {
            var paths = container.pathItems;
            for (i = 0; i < paths.length; i++) {
                try {
                    if (paths[i].clipping === true && isDirectChild(paths[i], container)) {
                        addUniqueReference(found, paths[i]);
                    }
                } catch (ignorePathClippingFlag) {}
            }
        } catch (ignorePathItems) {}
        try {
            var compounds = container.compoundPathItems;
            for (i = 0; i < compounds.length; i++) {
                if (!isDirectChild(compounds[i], container)) { continue; }
                var compoundPaths = compounds[i].pathItems;
                var j;
                for (j = 0; j < compoundPaths.length; j++) {
                    try {
                        if (compoundPaths[j].clipping === true) { addUniqueReference(found, compoundPaths[j]); }
                    } catch (ignoreCompoundClippingFlag) {}
                }
            }
        } catch (ignoreCompoundPathItems) {}
    }

    function getGeometricBounds(item) {
        var b;
        try { b = item.geometricBounds; } catch (geometricError) { b = null; }
        if (!b || b.length < 4) { return getBounds(item); }
        var result = [Number(b[0]), Number(b[1]), Number(b[2]), Number(b[3])];
        if (!isFinite(result[0]) || !isFinite(result[1]) || !isFinite(result[2]) || !isFinite(result[3]) ||
            result[2] < result[0] || result[1] < result[3]) {
            return getBounds(item);
        }
        return result;
    }

    function getBounds(item) {
        var b;
        try { b = item.visibleBounds; } catch (visibleError) { b = null; }
        if (!b || b.length < 4) {
            try { b = item.geometricBounds; } catch (geometricError) { b = null; }
        }
        if (!b || b.length < 4) {
            throw new Error("An item's bounds could not be read.");
        }
        var result = [Number(b[0]), Number(b[1]), Number(b[2]), Number(b[3])];
        if (!isFinite(result[0]) || !isFinite(result[1]) || !isFinite(result[2]) || !isFinite(result[3])) {
            throw new Error("An item's bounds contain invalid coordinates.");
        }
        if (result[2] < result[0] || result[1] < result[3]) {
            throw new Error("An item's bounds have an invalid width or height.");
        }
        return result;
    }

    function translateUsingBounds(item, oldBounds, targetLeft, targetTop) {
        var dx = targetLeft - oldBounds[0];
        var dy = targetTop - oldBounds[1];
        if (dx !== 0 || dy !== 0) {
            item.translate(dx, dy);
        }
    }

    function getExpandedArtboardRect(originalRect, items, extraItem) {
        var left = originalRect[0];
        var top = originalRect[1];
        var right = originalRect[2];
        var bottom = originalRect[3];
        var allItems = [];
        var i;
        for (i = 0; i < items.length; i++) {
            if (items[i]) { allItems.push(items[i]); }
        }
        if (extraItem) { allItems.push(extraItem); }

        for (i = 0; i < allItems.length; i++) {
            var b = getArtworkBounds(allItems[i]);
            if (b[0] < left) { left = b[0]; }
            if (b[1] > top) { top = b[1]; }
            if (b[2] > right) { right = b[2]; }
            if (b[3] < bottom) { bottom = b[3]; }
        }
        return [left, top, right, bottom];
    }

    function hasItemsOutsideRect(items, rect) {
        var tolerance = 0.5;
        var i;
        for (i = 0; i < items.length; i++) {
            if (!items[i]) { continue; }
            var b = getArtworkBounds(items[i]);
            if (b[0] < rect[0] - tolerance || b[1] > rect[1] + tolerance ||
                b[2] > rect[2] + tolerance || b[3] < rect[3] - tolerance) {
                return true;
            }
        }
        return false;
    }

    function copyRect(rect) {
        return [Number(rect[0]), Number(rect[1]), Number(rect[2]), Number(rect[3])];
    }

    function rectsEqual(a, b) {
        var i;
        for (i = 0; i < 4; i++) {
            if (Math.abs(a[i] - b[i]) > 0.01) { return false; }
        }
        return true;
    }

    function removeCreatedItems(items) {
        var i;
        for (i = items.length - 1; i >= 0; i--) {
            try { items[i].remove(); } catch (ignoreRemoveItem) {}
        }
    }

    function removeCreatedArtboards(documentRef, indices) {
        var i;
        for (i = indices.length - 1; i >= 0; i--) {
            try {
                if (documentRef.artboards.length > 1) {
                    documentRef.artboards.remove(indices[i]);
                }
            } catch (ignoreRemoveArtboard) {}
        }
    }

    function selectOutputItems(documentRef, items) {
        try { documentRef.selection = []; } catch (clearSelectionError) {}
        var i;
        for (i = 0; i < items.length; i++) {
            try { items[i].selected = true; } catch (ignoreSelectItem) {}
        }
    }

    // Parse one-based row/variable lists such as "2", "2-15", or "2,4,7-9".
    // Returned numbers are sorted, unique, and validated against the current item count.
    function parseIndexSpecification(spec, maximum, itemLabel) {
        var parts = String(spec).split(",");
        var indexes = [];
        var seen = {};
        var i;
        for (i = 0; i < parts.length; i++) {
            var token = trim(parts[i]);
            if (!token) { throw new Error("An item in the " + itemLabel + " list is empty."); }
            var start;
            var end;
            var rangeMatch = token.match(/^(\d+)\s*-\s*(\d+)$/);
            if (rangeMatch) {
                start = parseInt(rangeMatch[1], 10);
                end = parseInt(rangeMatch[2], 10);
                if (start > end) { throw new Error("Invalid range '" + token + "'. The first number must not exceed the last."); }
            } else if (/^\d+$/.test(token)) {
                start = parseInt(token, 10);
                end = start;
            } else {
                throw new Error("Invalid item '" + token + "'. Use numbers or ranges separated by commas, such as 2-5,7,9.");
            }
            if (start < 1 || end > maximum) {
                throw new Error(itemLabel + " numbers must be between 1 and " + maximum + ". Invalid item: '" + token + "'.");
            }
            var n;
            for (n = start; n <= end; n++) {
                if (!seen[String(n)]) { seen[String(n)] = true; indexes.push(n); }
            }
        }
        indexes.sort(function (a, b) { return a - b; });
        return indexes;
    }

    // Returns a new array after removing the 1-based indices parsed from a list such as
    // "2", "2-15", or "2,4". The original list and its row objects are not mutated.
    function filterItemsByIndexSpecification(items, specification, itemLabel) {
        if (!items || typeof items.length !== "number") { throw new Error("The list to filter is invalid."); }
        var removed = parseIndexSpecification(specification, items.length, itemLabel || "Item");
        var removeMap = {};
        var i;
        for (i = 0; i < removed.length; i++) { removeMap[String(removed[i])] = true; }
        var retained = [];
        for (i = 0; i < items.length; i++) {
            if (!removeMap[String(i + 1)]) { retained.push(items[i]); }
        }
        return retained;
    }

    // ----------------------- Validation helpers -----------------------
    function validateConfiguration(config, recordCount) {
        if (recordCount < 1) {
            throw new Error("At least one record is required.");
        }
        if (config.layout === "Grid" || (config.layout === "Separate Artboards" && (config.artboardArrangement || "Grid") === "Grid")) {
            if (config.columns * config.rows < recordCount) {
                throw new Error(
                    config.layout + " has capacity for only " + (config.columns * config.rows) + " record(s), but " + recordCount + " were entered.\n\n" +
                    "Increase Columns or Rows so Columns x Rows is at least the record count."
                );
            }
        }
        if (config.layout === "Same Position" && recordCount > 1) {
            // This is valid, but intentionally overlapping. A confirmation is shown before generation.
        }
    }

    function parsePositiveInteger(value, label) {
        var n = parseInt(trim(value), 10);
        if (!isFinite(n) || n < 1 || String(n) !== trim(value)) {
            throw new Error(label + " must be a positive whole number.");
        }
        return n;
    }

    function parseNonNegativeNumber(value, label) {
        var s = trim(value).replace(/,/g, ".");
        if (s === "") {
            throw new Error(label + " is required. Enter 0 or a positive number in centimeters.");
        }
        var n = Number(s);
        if (!isFinite(n) || n < 0) {
            throw new Error(label + " must be 0 or a positive number in centimeters.");
        }
        return n;
    }

    function cmToPoints(cm) {
        return cm * PT_PER_CM;
    }

    function makeOutputName(number) {
        var digits = String(number);
        while (digits.length < 3) { digits = "0" + digits; }
        return "Output_" + digits;
    }

    function errorMessage(error) {
        try {
            if (error && error.message) { return String(error.message); }
            return String(error);
        } catch (ignoreErrorMessage) {
            return "An unknown error occurred.";
        }
    }

    function trim(value) {
        return String(value).replace(/^\s+|\s+$/g, "");
    }

    function contains(array, value) {
        var i;
        for (i = 0; i < array.length; i++) {
            if (array[i] === value) { return true; }
        }
        return false;
    }

    function addUniqueReference(array, value) {
        var i;
        for (i = 0; i < array.length; i++) {
            if (array[i] === value) { return; }
        }
        array.push(value);
    }

}());
