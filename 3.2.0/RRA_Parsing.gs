/*
 * @NotOnlyCurrentDoc
 */

/* This Source Code Form is subject to the terms of the Mozilla Public
* License, v. 2.0. If a copy of the MPL was not distributed with this
* file, You can obtain one at https://mozilla.org/MPL/2.0/. */

/* This can be used from a spreadsheet or externally to process RRA GDocs and store them in a spreadsheet or db-like
 * format
 */
const levels_map = {"UNKNOWN": -1,
                      "LOW": 0,
                      "MEDIUM": 1,
                      "HIGH": 2,
                      "MAXIMUM": 3}

const levels = ['LOW', 'MEDIUM', 'HIGH', 'MAXIMUM']; //order matters
const classification_map = {"UNKNOWN": -1,
                      "Public": 0,
                      "Internal Confidential": 1,
                      "Specific Workgroups Only": 2,
                      "Specific Individuals Only": 3}
const classifications = ['Public', 'Internal Confidential', 'Specific Workgroups Only', 'Specific Individuals Only'];

function refreshRRA() {
  process_all_rra_docs();
}

function process_all_rra_docs() {
  var driveid = 'YOUR-DRIVE-ID-HERE'; // Where assessments are stored, this is a drive id which can be extracted from
  // your Google Drive URL
  // the templates, so that we don't process that
  var template_skip_ids = [
    'YOUR-TEMPLATE-ID-HERE'
  ];
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = s.getSheetByName('RRA3');
  var folder = DriveApp.getFolderById(driveid);
  var files = folder.getFiles();
  
  //Start fresh
  sheet.clearContents();
  // Headers
  sheet.appendRow(['Link', 'Name', 'Team', 'Reviewers', 'Main Data Classification', 'Highest Risk Impact', 'Shortlink', 'Recommendations', 'Highest Recommendation', 'Creation date', 'Modification date', 'Created by']);

  while (files.hasNext()) {
    var file = files.next();
    var exit_loop = false;
    // Skip templates
    for (var i in template_skip_ids) {
      if (file.getId() == template_skip_ids[i]) {
        exit_loop=true;
      }
    }
    if (exit_loop) {
      continue;
    }

    // Skip non-Documents and erroneous files
    try {
      var doc = DocumentApp.openById(file.getId());
    } catch (e) {
      console.warn("Invalid file, skipping: "+e)
      continue;
    }

    var rra_name = clean_fname(file.getName());
    s.toast("Importing RRA: "+rra_name+"...");
    console.log("Importing RRA: "+rra_name+" id: "+file.getId());

    // Remove footer with errors if any
    try {
      var footer = doc.getFooter();
      footer.clear();
      footer.removeFromParent();
    } catch (e) {        
    }

    // Import
    try {
      var results = null;

      if (file.getDateCreated() >= new Date("2020-07-01")) {
        // RRAs created from template 1.3, created after this date
        // (XXX whenever AppScript allows to query version, use the version instead of the date!)
        results = import_rra_13(rra_name, doc, file.getId());
      } else {
        results = import_rra(rra_name, doc, file.getId());
      }
    } catch (e) {
      console.warn("Failed to parse RRA: "+e);
      // Inform user of error
      // XXX this might be better if it used a comment instead of the footer
      var footer = doc.addFooter();
      var txt = footer.setText("This RRA failed to parse correctly. Please check the format and ensure that you follow the template.").editAsText();
      txt.setBold(true);
      txt.setBackgroundColor("#FF0000");
    }
    if(results != null) {
      insert_rra('https://docs.google.com/document/d/'+file.getId(), rra_name, sheet, results);
    }
  }
  s.toast("All done!");
}

// Import RRA doc to register
// v1.3
function import_rra_13(rra_name, doc, fid) {
  var docid = doc.getId();
  var tables = doc.getBody().getTables();
  var paragraphs = doc.getBody().getParagraphs();
  // These are not the manual "marked as reviewed" dates, but actual modification/creation date of the document
  var creation_date = DriveApp.getFileById(fid).getDateCreated();
  var modification_date = DriveApp.getFileById(fid).getLastUpdated();

  var recs = [];
  var highest_rec_rank = 0;
  var results = [];
  
  // First table has metadata
  var meta_table = tables[0];
  var celldata = "";
  var cellhdr = "";
  var cells = {};
  var valid_cellhdr = ["Risk owner", "Team", "Reviewers", "Main Data Classification", "Highest Risk Impact", "Shortlink"];
  
  for (var i=0;i<meta_table.getNumRows();i++) {
    cellhdr = meta_table.getCell(i, 0).getText().split('\n')[0];
    celldata = meta_table.getCell(i,1).getText().split('\n')[0];
    if (valid_cellhdr.indexOf(cellhdr) !== -1) {
      cells[cellhdr] = celldata;
    }
  }
  // Has to be in the right sheet row order
  if (cells["Team"] !== undefined) {
    results.push(["Team", cells["Team"]]);
  } else {
    results.push(["Teams", cells["Risk owner"]]);
  }
  results.push(["Reviewers", cells["Reviewers"]]);
  results.push(["Main Data Classification", cells["Main Data Classification"]]);
  results.push(["Highest Risk Impact", cells["Highest Risk Impact"]]);
  results.push(["Shortlink", cells["Shortlink"]]);
  
  // Find recommendations (this loop is a little hackish, but i couldnt find a good way to iterate without changing the original docs/adding bookmarks f.e.)
  for (var p=0;p<paragraphs.length;p++) {
    var current = paragraphs[p];
    if (current.getText() == 'Recommendations') {
      for (var p1=p;p1<paragraphs.length;p1++) {
        var line = paragraphs[p1];
        if (line.getType() == DocumentApp.ElementType.LIST_ITEM) {
          //Find if we have a recommendation level associated with the recommendation list item
          var rec_level = 'UNKNOWN';
          for (var l=0;l<levels.length;l++) {
            if (line.getText().split(' ')[0] == levels[l]) {
              rec_level = levels[l];
              // Find highest rec
              if (l > highest_rec_rank) {
                highest_rec_rank = l;
              }
              break;
            }
          }
          // Check that the rec is not solved/strikedout
          // And that it isn't just a sub list item (i.e. it has a rec_level)
          var attrs = line.getAttributes();
          if (attrs[DocumentApp.Attribute.STRIKETHROUGH] != true && rec_level != 'UNKNOWN') {
            recs.push([highest_rec_rank, rec_level, line.getText()]);
          } else {
            //Noisy and there is no "debug" level
            //Logger.log('Recommendation already striked out or missing level, skipping');
          }
        }
      }
      break;
    }
  }
  results.push(['Recommendations', recs.length]);
  results.push(['Highest Recommendation', levels[highest_rec_rank]]);
  results.push(['Creation date', creation_date]);
  results.push(['Modification date', modification_date]);

  // getOwner() does not work in this scope, so we hack.
  var created_by = "unknown";
  var reviewers = cells["Reviewers"];
  try {
    if (reviewers.indexOf(' ')) {
      created_by = reviewers.split(' ')[0];
    }
    if (reviewers.indexOf('@')) {
      created_by = created_by.split('@')[0];
    }
    if (reviewers.indexOf(',')) {
      created_by = created_by.split(',')[0];
    }
  } catch(e) {
    console.warn("Unparsable reviewer: "+e)
  }
  results.push(['Created by', created_by]);
  return results
}


// v1.2
function import_rra(rra_name, doc, fid) {
  var docid = doc.getId();
  var tables = doc.getBody().getTables();
  var paragraphs = doc.getBody().getParagraphs();
  // These are not the manual "marked as reviewed" dates, but actual modification/creation date of the document
  var creation_date = DriveApp.getFileById(fid).getDateCreated();
  var modification_date = DriveApp.getFileById(fid).getLastUpdated();

  var recs = [];
  var highest_rec_rank = 0;
  var results = [];
  
  // First table has metadata
  var meta_table = tables[0];
  var celldata = "";
  var cellhdr = "";
  var cells = {};
  var valid_cellhdr = ["Risk owner", "Team", "Reviewers", "Main Data Classification", "Highest Risk Impact", "Shortlink"];
  
  for (var i=0;i<meta_table.getNumRows();i++) {
    cellhdr = meta_table.getCell(i, 0).getText().split('\n')[0];
    celldata = meta_table.getCell(i,1).getText().split('\n')[0];
    if (valid_cellhdr.indexOf(cellhdr) !== -1) {
      cells[cellhdr] = celldata;
    }
  }
  // Has to be in the right sheet row order
  if (cells["Team"] !== undefined) {
    results.push(["Team", cells["Team"]]);
  } else {
    results.push(["Teams", cells["Risk owner"]]);
  }
  results.push(["Reviewers", cells["Reviewers"]]);
  results.push(["Main Data Classification", cells["Main Data Classification"]]);
  results.push(["Highest Risk Impact", cells["Highest Risk Impact"]]);
  results.push(["Shortlink", cells["Shortlink"]]);
  
  // Find recommendations (this loop is a little hackish, but i couldnt find a good way to iterate without changing the original docs/adding bookmarks f.e.)
  for (var p=0;p<paragraphs.length;p++) {
    var current = paragraphs[p];
    if (current.getText() == 'Recommendations') {
      for (var p1=p;p1<paragraphs.length;p1++) {
        var line = paragraphs[p1];
        if (line.getType() == DocumentApp.ElementType.LIST_ITEM) {
          //Find if we have a recommendation level associated with the recommendation list item
          var rec_level = 'UNKNOWN';
          for (var l=0;l<levels.length;l++) {
            if (line.getText().split(' ')[0] == levels[l]) {
              rec_level = levels[l];
              // Find highest rec
              if (l > highest_rec_rank) {
                highest_rec_rank = l;
              }
              break;
            }
          }
          // Check that the rec is not solved/strikedout
          // And that it isn't just a sub list item (i.e. it has a rec_level)
          var attrs = line.getAttributes();
          if (attrs[DocumentApp.Attribute.STRIKETHROUGH] != true && rec_level != 'UNKNOWN') {
            recs.push([highest_rec_rank, rec_level, line.getText()]);
          } else {
            //Noisy and there is no "debug" level
            //Logger.log('Recommendation already striked out or missing level, skipping');
          }
        }
      }
      break;
    }
  }
  results.push(['Recommendations', recs.length]);
  results.push(['Highest Recommendation', levels[highest_rec_rank]]);
  results.push(['Creation date', creation_date]);
  results.push(['Modification date', modification_date]);
  // getOwner() does not work in this scope, so we hack.
  var created_by = "unknown";
  var reviewers = cells["Reviewers"];
  try {
    if (reviewers.indexOf(' ')) {
      created_by = reviewers.split(' ')[0];
    }
    if (reviewers.indexOf('@')) {
      created_by = created_by.split('@')[0];
    }
    if (reviewers.indexOf(',')) {
      created_by = created_by.split(',')[0];
    }
  } catch(e) {
    console.warn("Unparsable reviewer: "+e)
  }
  results.push(['Created by', created_by]);
  return results
}

/*
 * Returns a numerical equivalent for a level and verify the syntax is OK
 */
function standardize(level) {
  var valid = false;
  var num_val = -1;

  //Fix some typoes..
  var tmp = "";
  for (var x=0; x< 5; x++) {
    tmp = level.split(' ')[x];
    if (tmp.length > 0) {
      level = tmp;
      break;
    }
  }

  for (var i in levels) {
    if (levels[i] === level) {
      valid=true;
      num_val = levels_map[level];
      break
    }
  }
 
  if (!valid) {
    console.error("Invalid level: "+level);
  }
  return num_val;
}

/* 
 * Same as standardize, but in text
 */
function standardize_txt(level) {
  return levels[standardize(level)];
}

/*
 * Returns a numerical equivalent for a classification and verify the syntax is OK
 */
function standardize2(classification) {
  var valid = false;
  var num_val = -1;

  if (classification === undefined || classification.length == 0) {
    console.error("Empty input to standardize2() will default to unknown classification");
    return num_val;
  }

  //Fix some typoes..
  var tmp = "";
  for (var x=0; x< 5; x++) {
    tmp = classification.split(' ')[x];
    if (tmp.length > 0) {
      classification = tmp;
      break;
    }
  }

  for (var i in classifications) {
    if (classifications[i] === classification) {
      valid=true;
      num_val = classification_map[classification];
      break
    }
  }
 
  if (!valid) {
    console.error("Invalid classification: "+classification);
  }
  return num_val;
}

/*
 * Same as standardize, but in text
 */
function standardize2_txt(classification) {
  return classifications[standardize2(classification)];
}

// clean up filename a bit
function clean_fname(fname) {
  var clean_name = fname.split(' - ')[1];
  if (clean_name === undefined) {
    clean_name = fname.split('RRA ')[1];
  }
  if (clean_name === undefined) {
    clean_name = fname;
  }
  return clean_name
}

// Insert results in register
function insert_rra(docid, fname, sheet, results) {
  var s = SpreadsheetApp.getActiveSpreadsheet()

  var row = [docid, fname];
  var valid = true;
  
  for (var y = 0; y < results.length; y++) {
    row.push(results[y][1]);
    if (results[y][1] == '') {
      valid = false;
    }
  }
  //Logger.log(row);
  sheet.appendRow(row);
  if (!valid) {
    console.warn("Row is missing elements: "+row);
  }
}
