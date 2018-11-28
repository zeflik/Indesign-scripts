/*  Created by Józef Niemiec / 20.11.2018
    email: kontakt@jozefniemiec.pl 
    
    Script removes oversets:
    - on selected items (if there is a selection)
    - on all page items in the active document (if no selection)
    Used algorithm:
    - text frames - try to adjust the height, then width, then both. If overset still exist rollback changes.
    - tables - set autoGrow = true
*/

const SCRIPT_NAME = "Remove all text Oversets";

app.doScript(
  main,
  ScriptLanguage.JAVASCRIPT, [],
  UndoModes.FAST_ENTIRE_SCRIPT,
  SCRIPT_NAME
);

function main() {
  if (app.documents.length == 0) {
    alert("No active document.");
    return;
  }

  if (app.activeDocument.selection.length > 0) {
    removeTextFramesOversets(app.activeDocument.selection);
    return;
  }

  if (!confirm("Remove all document oversets?", false)) {
    return;
  };


  try {
    var tablesCells = app.activeDocument.stories.everyItem().tables.everyItem().cells.everyItem().getElements();
    removeTablesOversets(tablesCells);
  } catch (e) {}
  var allPageItems = app.activeDocument.allPageItems;
  removeTextFramesOversets(allPageItems);
  alert("Done");
}

function removeTablesOversets(tablesCells) {
  for (i = 0; i < tablesCells.length; i++) {
    if (tablesCells[i].insertionPoints.length > 0 && tablesCells[i].insertionPoints.lastItem().parentTextFrames.length == 0) {
      tablesCells[i].autoGrow = true;
    }
  }
}

function removeTextFramesOversets(allPageItems) {
  var pageItem, oldSizingReferencePoint, oldAutoSizingType, oldGeometricBounds;
  for (i = 0; i < allPageItems.length; i++) {
    pageItem = allPageItems[i];
    if (pageItem == "[object TextFrame]" && pageItem.overflows) {
      oldSizingReferencePoint = pageItem.textFramePreferences.autoSizingReferencePoint;
      oldAutoSizingType = pageItem.textFramePreferences.autoSizingType;
      oldGeometricBounds = pageItem.geometricBounds;
      pageItem.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_LEFT_POINT;
      pageItem.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
      if (pageItem.overflows) {
        pageItem.textFramePreferences.autoSizingType = AutoSizingTypeEnum.WIDTH_ONLY;
      }
      if (pageItem.overflows) {
        pageItem.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_AND_WIDTH;
      }
      pageItem.textFramePreferences.autoSizingReferencePoint = oldSizingReferencePoint;
      pageItem.textFramePreferences.autoSizingType = oldAutoSizingType;
      if (pageItem.overflows) {
        pageItem.geometricBounds = oldGeometricBounds;
      }
    }
  }
}
