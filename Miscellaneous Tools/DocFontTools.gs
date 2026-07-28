/**
 * Title: Google Docs Font Inspector & Rogue Font Highlighter
 * Description: Google Docs Apps Script tool that adds a custom "Font Tools" menu to Google Docs.
 *              1. Highlight Non-Brand Fonts: Scans document text attributes and highlights non-brand fonts in yellow.
 *              2. Generate Font Report: Extracts font, size, and styling metrics into a new Google Sheet report.
 * Author: David Knuffke
 * Target: Google Docs / Apps Script
 */

function onOpen() {
  DocumentApp.getUi().createMenu('Font Tools')
      .addItem('Generate Font Report', 'generateDocFontReport')
      .addItem('Highlight Non-Brand Fonts', 'highlightRogueFonts') // New menu item
      .addToUi();
}

// ---------------------------------------------------------
// NEW FUNCTION: Highlight fonts that don't match your brand
// ---------------------------------------------------------
function highlightRogueFonts() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();
  
  // --- SET YOUR BRAND FONT HERE ---
  // Must exactly match the font name as it appears in Google Docs (e.g., "Proxima Nova", "Open Sans")
  var BRAND_FONT = "Arial"; 
  var HIGHLIGHT_COLOR = "#FFFF00"; // Yellow hex code
  
  var rogueCount = 0;

  for (var p = 0; p < paragraphs.length; p++) {
    var textObj = paragraphs[p].editAsText();
    var textString = textObj.getText();
    
    if (textString.trim().length === 0) continue; // Skip empty paragraphs
    
    // Find every index where the text styling changes
    var indices = textObj.getTextAttributeIndices();
    
    for (var i = 0; i < indices.length; i++) {
      var start = indices[i];
      var end = (i + 1 < indices.length) ? indices[i + 1] : textString.length;
      
      var font = textObj.getFontFamily(start);
      
      // In Google Docs, if a font returns 'null', it means it is inheriting the default 'Normal text' theme font.
      // We only flag fonts that are explicitly set and do NOT match the brand font.
      if (font !== null && font !== BRAND_FONT) {
        if (end > start) {
          // highlight the rogue text (end offset is inclusive, so we subtract 1)
          textObj.setBackgroundColor(start, end - 1, HIGHLIGHT_COLOR);
          rogueCount++;
        }
      }
    }
  }
  
  DocumentApp.getUi().alert(
    'Scan Complete', 
    'Highlighted ' + rogueCount + ' sections of text that do not match your brand font (' + BRAND_FONT + ').', 
    DocumentApp.getUi().ButtonSet.OK
  );
}

// ---------------------------------------------------------
// ORIGINAL FUNCTION: Generate the Spreadsheet Report
// ---------------------------------------------------------
function generateDocFontReport() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();
  
  var reportData = [["Location (Paragraph #)", "Text Snippet", "Font Family", "Font Size", "Styles"]];
  
  for (var p = 0; p < paragraphs.length; p++) {
    var textObj = paragraphs[p].editAsText();
    var textString = textObj.getText();
    
    if (textString.trim().length === 0) continue; 
    
    var indices = textObj.getTextAttributeIndices();
    
    for (var i = 0; i < indices.length; i++) {
      var start = indices[i];
      var end = (i + 1 < indices.length) ? indices[i + 1] : textString.length;
      
      var font = textObj.getFontFamily(start) || "Default/Theme";
      var size = textObj.getFontSize(start) || "Default/Theme";
      
      var bold = textObj.isBold(start) ? "Bold " : "";
      var italic = textObj.isItalic(start) ? "Italic " : "";
      var underline = textObj.isUnderline(start) ? "Underline " : "";
      var styles = (bold + italic + underline).trim() || "Regular";
      
      var snippet = textString.substring(start, Math.min(start + 40, end)).replace(/\n/g, "");
      if (snippet.trim() === "") continue;
      
      reportData.push([p + 1, snippet, font, size, styles]);
    }
  }
  
  var sheet = SpreadsheetApp.create("Doc Font Report: " + doc.getName());
  sheet.getActiveSheet().getRange(1, 1, reportData.length, reportData[0].length).setValues(reportData);
  
  DocumentApp.getUi().alert(
    'Report Complete', 
    'A new spreadsheet named "Doc Font Report: ' + doc.getName() + '" has been created in your Google Drive root folder.', 
    DocumentApp.getUi().ButtonSet.OK
  );
}
