/**
 * @OnlyCurrentDoc
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 */

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} ev The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(ev) {
  onOpen(ev);
}

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} ev The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(ev) {
  var ui = DocumentApp.getUi();
  // Inside add-on context
  var menu1 = ui.createAddonMenu()
    .addSubMenu(ui.createMenu('Insert Data Classification Label')
                    .addItem('Public', 'class1')
                    .addItem('Internal Confidential', 'class2')
                    .addItem('Specific Workgroups only', 'class3')
                    .addItem('Specific Individuals only', 'class4')
                    .addSeparator()
                    .addItem('Business Data', 'class01')
                    .addItem('User Data', 'class02')
                    .addItem('Identifiable Information', 'class03')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('Insert Level')
                    .addItem('LOW', 'risk1')
                    .addItem('MEDIUM', 'risk2')
                    .addItem('HIGH ', 'risk3')
                    .addItem('MAXIMUM', 'risk4')
    )
    .addSeparator()
    .addItem("Show Risk Levels reference", 'sidebar_help')

// Outside add-on context
  var menu2 = ui.createMenu('RRA Utilities')
    .addSubMenu(ui.createMenu('Insert Data Classification Label')
                    .addItem('Public', 'class1')
                    .addItem('Confidential', 'class2')
                    .addItem('Specific Workgroups Only', 'class3')
                    .addItem('Specific Individuals Only', 'class4')
                    .addSeparator()
                    .addItem('Business Data', 'class01')
                    .addItem('User Data', 'class02')
                    .addItem('Identifiable Information', 'class03')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('Insert Level')
                    .addItem('LOW', 'risk1')
                    .addItem('MEDIUM', 'risk2')
                    .addItem('HIGH ', 'risk3')
                    .addItem('MAXIMUM', 'risk4')
    )
    .addSeparator()
    .addItem("Show Risk Levels reference", 'sidebar_help')

  menu1.addToUi();
  menu2.addToUi();
}

function sidebar_help() {
  html = HtmlService.createHtmlOutputFromFile("guide.html")
      .setTitle('Risk Levels')
      .setWidth(600);
  DocumentApp.getUi()
      .showSidebar(html);
}

function insertFormattedText(title1, color1) {
  if (color1 == null) {
    color1 = "#000000";
  }
  var doc = DocumentApp.getActiveDocument();
  var c; // cursor

  // If text is selected, remove it first
  var sel = DocumentApp.getActiveDocument().getSelection();
  if (sel) {
    var r = sel.getRangeElements()[0];
    var pos_e = r.getEndOffsetInclusive();
    var pos_s = r.getStartOffset();
    r.getElement().asText().deleteText(pos_s, pos_e);

    var p = doc.newPosition(r.getElement(), r.getStartOffset());
    doc.setCursor(p);
    c = doc.getCursor();
  } else {
    // Reposition cursor at begining of line for convenience
    c = doc.getCursor();
    var p = doc.newPosition(c.getElement(), 0);
    // this inserts text in front of the cursor, so we end up with a space after the insertion. But this is also here because of a bug in AppScript where c.insertText will silently fail or cause a server error (this happens when the cursor is moved from end of the line to start of the line)
    p.insertText(" "); 
    doc.setCursor(p);
    c = doc.getCursor();

  }
  if (c) {
    var style = {};

    style[DocumentApp.Attribute.FOREGROUND_COLOR] = color1;
    style[DocumentApp.Attribute.FONT_FAMILY] = 'Google Sans';
    style[DocumentApp.Attribute.FONT_SIZE] = 11;
    style[DocumentApp.Attribute.UNDERLINE] = false;
    style[DocumentApp.Attribute.BOLD] = true;

    c.insertText(title1);

    var txt = c.getSurroundingText();
    var cpos_start = c.getSurroundingTextOffset();
    var cpos_end = cpos_start + title1.length-1;

    txt.setAttributes(cpos_start, cpos_end, style);
  }
}

function class1() {
  insertFormattedText('Public', '#7a7a7a');
}
function class2() {
  insertFormattedText('Internal Confidential', '#4a6785');
}
function class3() {
  insertFormattedText('Specific Workgroups Only', '#d04437');
}
function class4() {
  insertFormattedText('Specific Individuals Only', '#d04437');
}
function class01() {
  insertFormattedText('Business Data');
}
function class02() {
  insertFormattedText('User Data');
}
function class03() {
  insertFormattedText('Identifiable Information');
}

function risk1() {
  insertFormattedText('LOW', '#7a7a7a');
}
function risk2() {
  insertFormattedText('MEDIUM', '#4a6785');
}
function risk3() {
  insertFormattedText('HIGH', '#ebbd30');
}
function risk4() {
  insertFormattedText('MAXIMUM', '#d04437');
}
